import os
import io
import sys
import time
import json
import html
import base64
import hashlib
import asyncio
import logging
import threading
import requests
from contextlib import asynccontextmanager
from datetime import datetime
from urllib.parse import quote
from urllib3.util.retry import Retry
from requests.adapters import HTTPAdapter
from dotenv import load_dotenv
from docx import Document
import openpyxl
import pdfplumber
from pptx import Presentation
from fastapi import Depends, FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, HTMLResponse
from pydantic import BaseModel, Field
import uvicorn
import db
import s3

load_dotenv()

if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

# Direct model cache to /tmp so it works on SAP BTP's Linux filesystem
os.environ.setdefault("SENTENCE_TRANSFORMERS_HOME", "/tmp/models")

from sentence_transformers import SentenceTransformer


# ── Logging ───────────────────────────────────────────────────────────────────

logging.basicConfig(
    level   = logging.INFO,
    format  = "%(asctime)s  %(levelname)-8s  %(name)s — %(message)s",
    datefmt = "%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("sharepoint_api")


# ── Config ────────────────────────────────────────────────────────────────────

_embed_model = None

def _get_embed_model():
    global _embed_model
    if _embed_model is None:
        _embed_model = SentenceTransformer("all-MiniLM-L6-v2")
    return _embed_model
_token_cache = {"token": None, "expires_at": 0}
_token_lock  = threading.Lock()

# Detect runtime environment
_ON_CF = bool(os.environ.get("VCAP_SERVICES"))

TENANT_ID     = os.getenv("TENANT_ID")
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
HOSTNAME      = os.getenv("HOSTNAME",      "z0y8z.sharepoint.com")
SITE_PATH     = os.getenv("SITE_PATH",     "/sites/BTP-CatalystPlatform")
DOWNLOAD_PATH = os.getenv("DOWNLOAD_PATH", "/tmp")          # SAP BTP: use /tmp or Object Store mount
_log_day = datetime.now().strftime("%Y%m%d")
os.makedirs(os.path.join(DOWNLOAD_PATH, _log_day), exist_ok=True)
_fh = logging.FileHandler(os.path.join(DOWNLOAD_PATH, _log_day, f"api_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"), encoding="utf-8")
_fh.setFormatter(logging.Formatter(
    "%(asctime)s  %(levelname)-8s  %(name)s — %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
))
logging.getLogger().addHandler(_fh)


def _call_logger(endpoint: str) -> logging.Logger:
    ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
    day     = ts[:8]
    log_dir = os.path.join(DOWNLOAD_PATH, day)
    os.makedirs(log_dir, exist_ok=True)
    path    = os.path.join(log_dir, f"{endpoint}_{ts}.log")
    logger  = logging.Logger(f"call.{endpoint}.{ts}")
    handler = logging.FileHandler(path, encoding="utf-8")
    handler.setFormatter(logging.Formatter(
        "%(asctime)s  %(levelname)-8s — %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    logger.propagate = False
    return logger


def _get_log_path(logger: logging.Logger) -> str | None:
    """Return the file path of the first FileHandler on a logger, or None."""
    for h in logger.handlers:
        if isinstance(h, logging.FileHandler):
            return h.baseFilename
    return None


GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPE      = "https://graph.microsoft.com/.default"
TOKEN_URL  = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
URL_BASE   = f"{GRAPH_BASE}/sites/{HOSTNAME}:{SITE_PATH}"

site_id  = None
drive_id = None


# ── Auth ──────────────────────────────────────────────────────────────────────

# BTP Destination Service caches
_dest_svc_creds     = None
_dest_svc_token     = {"token": None, "expires_at": 0}
_dest_svc_lock      = threading.Lock()
DESTINATION_NAME    = "BTP-CP_MS_GRAPH"


def _get_dest_svc_creds() -> dict:
    """Extract the BTP Destination Service binding from VCAP_SERVICES."""
    global _dest_svc_creds
    if _dest_svc_creds:
        return _dest_svc_creds
    vcap = json.loads(os.environ.get("VCAP_SERVICES", "{}"))
    for key in ("destination", "destinations"):
        if key in vcap:
            _dest_svc_creds = vcap[key][0]["credentials"]
            log.info("BTP Destination Service binding found")
            return _dest_svc_creds
    raise RuntimeError(
        "Destination service not bound — add 'destination' under services: in manifest.yml"
    )


def _get_dest_svc_token() -> str:
    """Get an OAuth token scoped to the BTP Destination Service API."""
    with _dest_svc_lock:
        if _dest_svc_token["token"] and time.time() < _dest_svc_token["expires_at"] - 60:
            return _dest_svc_token["token"]
        creds = _get_dest_svc_creds()
        resp  = requests.post(
            f"{creds['url']}/oauth/token",
            data={
                "grant_type":    "client_credentials",
                "client_id":     creds["clientid"],
                "client_secret": creds["clientsecret"],
            },
            timeout=15,
        )
        resp.raise_for_status()
        data = resp.json()
        _dest_svc_token["token"]      = data["access_token"]
        _dest_svc_token["expires_at"] = time.time() + data.get("expires_in", 3600)
        log.info("BTP Destination Service token refreshed")
        return _dest_svc_token["token"]


def _get_ms_graph_token_from_destination() -> tuple[str, int]:
    """
    Fetch an MS Graph access token via the BTP Destination Service.
    Returns (token, expires_in_seconds).
    Prefers the pre-fetched authToken from the Destination Service;
    falls back to calling Azure AD directly using the destination's credentials.
    """
    creds      = _get_dest_svc_creds()
    dest_token = _get_dest_svc_token()
    resp = requests.get(
        f"{creds['uri']}/destination-configuration/v1/destinations/{DESTINATION_NAME}",
        headers={"Authorization": f"Bearer {dest_token}"},
        timeout=15,
    )
    resp.raise_for_status()
    payload = resp.json()

    # Option A — Destination Service pre-fetched a Bearer token
    for auth in payload.get("authTokens", []):
        if auth.get("type", "").lower() == "bearer" and auth.get("value"):
            expires_in = int(auth.get("expires_in", 3600))
            log.info(f"MS Graph token obtained from Destination Service authToken (expires_in={expires_in}s)")
            return auth["value"], expires_in

    # Option B — extract credentials and call Azure AD directly
    config     = payload.get("destinationConfiguration", {})
    token_url  = config.get("tokenServiceURL") or config.get("TokenServiceURL", "")
    client_id  = config.get("clientId")        or config.get("ClientId", "")
    client_sec = config.get("clientSecret")    or config.get("ClientSecret", "")
    scope      = config.get("scope",  SCOPE)
    if not (token_url and client_id and client_sec):
        raise RuntimeError(
            f"Destination '{DESTINATION_NAME}' missing tokenServiceURL / clientId / clientSecret"
        )
    resp2 = requests.post(
        token_url,
        data={
            "grant_type":    "client_credentials",
            "client_id":     client_id,
            "client_secret": client_sec,
            "scope":         scope,
        },
        timeout=15,
    )
    resp2.raise_for_status()
    data2 = resp2.json()
    expires_in = int(data2.get("expires_in", 3600))
    log.info(f"MS Graph token obtained via destination credentials (expires_in={expires_in}s)")
    return data2["access_token"], expires_in


def _getAccessToken() -> str:
    with _token_lock:
        if _token_cache["token"] and time.time() < _token_cache["expires_at"] - 60:
            return _token_cache["token"]

        if _ON_CF:
            token, expires_in = _get_ms_graph_token_from_destination()
        else:
            # Local dev — use env vars
            resp = requests.post(TOKEN_URL, data={
                "grant_type":    "client_credentials",
                "client_id":     CLIENT_ID,
                "client_secret": CLIENT_SECRET,
                "scope":         SCOPE,
            })
            resp.raise_for_status()
            data       = resp.json()
            token      = data["access_token"]
            expires_in = data.get("expires_in", 3600)

        _token_cache["token"]      = token
        _token_cache["expires_at"] = time.time() + expires_in
        return _token_cache["token"]

def _headers() -> dict:
    return {"Authorization": f"Bearer {_getAccessToken()}", "Accept": "application/json"}


# ── Site and Drive ────────────────────────────────────────────────────────────

def _get_SiteDrive_ID() -> tuple:
    r = requests.get(URL_BASE, headers=_headers(), timeout=15)
    r.raise_for_status()
    siteid = r.json()["id"]
    r.close()
    d = requests.get(f"{GRAPH_BASE}/sites/{siteid}/drives", headers=_headers(), timeout=15)
    d.raise_for_status()
    drives = d.json().get("value", [])
    if not drives:
        raise RuntimeError("No document libraries found in SharePoint site")
    driveid = drives[0]["id"]
    d.close()
    return siteid, driveid


# ── Formatters ────────────────────────────────────────────────────────────────

def _fmt_size(size_bytes: int) -> str:
    if size_bytes is None:
        return "—"
    for unit in ("B", "KB", "MB", "GB"):
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


# ── Flatten helpers ───────────────────────────────────────────────────────────

def _flatten_item(item: dict) -> dict:
    ref      = item.get("parentReference", {})
    raw_path = ref.get("path", "")
    path     = raw_path.split(":")

    item_id     = item.get("id", "")
    item_name   = item.get("name", "")
    folder_path = path[1] if len(path) > 1 else ""
    dPath       = "/".join(quote(seg, safe="") for seg in f"{folder_path}/{item_name}".split("/"))
    basePath    = f"{GRAPH_BASE}{path[0]}:{dPath}:/content"

    return {
        "id":                 item_id,
        "name":               item_name,
        "type":               "Folder" if "folder" in item else item.get("file", {}).get("mimeType", "file"),
        "extension":          item.get("file", {}).get("fileExtension"),
        "size":               _fmt_size(item.get("size")),
        "size_bytes":         item.get("size"),
        "last_modified_date": item.get("lastModifiedDateTime", ""),
        "created_date":       item.get("createdDateTime", ""),
        "created_by":         item.get("createdBy", {}).get("user", {}).get("displayName", ""),
        "email_created_by":   item.get("createdBy", {}).get("user", {}).get("email", ""),
        "modified_by":        item.get("lastModifiedBy", {}).get("user", {}).get("displayName", ""),
        "email_modified_by":  item.get("lastModifiedBy", {}).get("user", {}).get("email", ""),
        "web_url":            item.get("webUrl", ""),
        "child_count":        item.get("folder", {}).get("childCount") if "folder" in item else None,
        "doc_path":           dPath,
        "versionURL":         f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/versions",
        "BaseURL":            basePath,
        "shared":             item.get("shared", {}).get("scope", ""),
        "quick_xor_hash":     item.get("file", {}).get("hashes", {}).get("quickXorHash"),
    }


def _flatten_versions(item: dict) -> dict:
    return {
        "version_id":           item.get("id"),
        "fileSize":             _fmt_size(item.get("size")),
        "size":                 item.get("size"),
        "lastModifiedDateTime": item.get("lastModifiedDateTime"),
        "lastmodified_by":      item.get("lastModifiedBy", {}).get("user", {}).get("email"),
        "download_url":         item.get("@microsoft.graph.downloadUrl"),
    }


# ── Session ───────────────────────────────────────────────────────────────────

def _build_session() -> requests.Session:
    session = requests.Session()
    retry   = Retry(
        total            = 3,
        backoff_factor   = 2,
        status_forcelist = [429, 500, 502, 503, 504],
    )
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.headers.update(_headers())
    return session


# ── Version fetcher ───────────────────────────────────────────────────────────

def _get_item_versions(id: str) -> list:
    rURL  = f"{GRAPH_BASE}/drives/{drive_id}/items/{id}/versions"
    rlist = []
    try:
        with _build_session() as session:
            res = session.get(rURL, timeout=(15, 120))
            res.raise_for_status()
            rlist = res.json().get("value", [])
            res.close()
            rlist = [_flatten_versions(x) for x in rlist]
            for x in rlist:
                x["item_id"]       = id
                x["version_count"] = len(rlist)
    except requests.exceptions.ReadTimeout:
        print(f"[TIMEOUT] item {id} — skipping")
    except requests.exceptions.HTTPError as e:
        sc = e.response.status_code if e.response is not None else "?"
        print(f"[HTTP ERROR] item {id} — {sc}")
    except Exception as e:
        print(f"[ERROR] item {id} — {e}")
    return rlist


# ── File listers ──────────────────────────────────────────────────────────────

def _get_root_files() -> list:
    log.info("Getting all root files from ~/root/delta")
    url    = f"{GRAPH_BASE}/drives/{drive_id}/root/delta"
    result = []
    first  = True
    while url:
        resp = requests.get(url, headers=_headers(), timeout=(15, 60))
        resp.raise_for_status()
        data = resp.json()
        resp.close()
        items = data.get("value", [])
        for f in (items[1:] if first else items):
            flat = _flatten_item(f)
            if flat["type"] != "Folder":
                result.append(flat)
        first = False
        url = data.get("@odata.nextLink")
    return result


def _get_all_files(src: str, use_delta: bool = True) -> list:
    endpoint = f"/root:{src}:/delta" if use_delta else f"/root:{src}"
    log.info(f"Getting all files from {endpoint}")
    url    = f"{GRAPH_BASE}/drives/{drive_id}{endpoint}"
    result = []
    while url:
        resp = requests.get(url, headers=_headers(), timeout=(15, 60))
        resp.raise_for_status()
        data = resp.json()
        resp.close()
        for f in data.get("value", []):
            flat = _flatten_item(f)
            if flat["type"] != "Folder":
                result.append(flat)
        url = data.get("@odata.nextLink")
    return result


def _get_item_by_id(item_id: str) -> dict:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
    r   = requests.get(url, headers=_headers(), timeout=15)
    r.raise_for_status()
    return _flatten_item(r.json())


def _search_files_by_name(query: str) -> list:
    url     = f"{GRAPH_BASE}/drives/{drive_id}/root/search(q='{query}')"
    results = []
    while url:
        r = requests.get(url, headers=_headers(), timeout=15)
        r.raise_for_status()
        data = r.json()
        for item in data.get("value", []):
            if "folder" not in item and query.lower() in item.get("name", "").lower():
                results.append(_flatten_item(item))
        url = data.get("@odata.nextLink")
    return results


def _get_root_files_with_versions(delay: float = 1.0) -> list:
    root_files = _get_root_files()
    total      = len(root_files)
    jdata      = []

    print(f"Found {total} files. Fetching versions...")
    log.info(f"Found {total} files. Fetching versions...")

    for i, x in enumerate(root_files, start=1):
        file_id  = x.get("id")
        doc_path = x.get("doc_path")
        print(f"  [{i}/{total}] {doc_path}")
        log.info(f"[{i}/{total}] {doc_path}")

        ver_list           = _get_item_versions(file_id)
        x["version_count"] = len(ver_list)
        x["versions_list"] = ver_list
        jdata.append(x)

        if i < total:
            time.sleep(delay)

    print(f"Done. {total} files processed.")
    log.info(f"Done. {total} files processed.")

    return jdata


# ── Version extractor ────────────────────────────────────────────────────────

def _split_versions(root_files: list) -> tuple:
    """
    Separates versions_list out of each file dict into a flat list.
    Annotates each version record with filename for traceability.
    version_count is kept in the file dict; versions_list is removed.
    Returns (clean_files, flat_versions).
    """
    clean_files   = []
    flat_versions = []
    for f in root_files:
        versions = f.pop("versions_list", None) or []
        filename = f.get("name", "")
        for v in versions:
            v["filename"] = filename
        flat_versions.extend(versions)
        clean_files.append(f)
    return clean_files, flat_versions


# ── JSON writers ──────────────────────────────────────────────────────────────

def _write_to_json(data: list, path: str = None) -> None:
    if path:
        with open(path, "a", encoding="utf-8") as fh:
            for record in data:
                fh.write(json.dumps(record, ensure_ascii=False) + "\n")
        print(f"\n{len(data)} records written to {path}")
        log.info(f"{len(data)} records written to {path}")


def _write_to_json_array_indent(data: list, path: str) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)
    print(f"\n{len(data)} records written to {path}")
    log.info(f"{len(data)} records written to {path}")


def _generate_db_analysis_html(data: dict) -> str:
    from datetime import datetime as _dt
    ts = _dt.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

    def fmt_bytes(b):
        if not b: return "—"
        for u in ("B","KB","MB","GB"):
            if b < 1024: return f"{b:.1f} {u}"
            b /= 1024
        return f"{b:.1f} TB"

    total_files         = data.get("total_files", 0)
    total_chunks        = data.get("total_chunks", 0)
    total_vectors       = data.get("total_vectors", 0)
    total_bytes         = data.get("total_size_bytes", 0)
    avg_chunk_chars     = data.get("avg_chunk_chars", 0)
    avg_chunks_per_file = round(total_chunks / total_files, 1) if total_files else 0
    vector_coverage     = round(total_vectors / total_chunks * 100, 1) if total_chunks else 0
    missing_cs          = data.get("files_missing_chunk_summary", 0)
    chunks_no_v         = data.get("chunks_without_vectors", 0)
    health_ok           = missing_cs == 0 and chunks_no_v == 0

    cs_card   = "card-green" if missing_cs  == 0 else "card-red"
    cs_col    = "green"      if missing_cs  == 0 else "orange"
    cs_sub    = "all files have a chunk summary record" if missing_cs == 0 else "in BRAIN_SP_FILES but absent from BRAIN_SP_FILES_CHUNK"
    vec_card  = "card-green" if chunks_no_v == 0 else "card-red"
    vec_col   = "green"      if chunks_no_v == 0 else "orange"
    vec_sub   = "full vector coverage" if chunks_no_v == 0 else "chunks with no matching BRAIN_VECTORS row"

    health_banner = (
        '<div style="background:#14532d33;border:1px solid #16a34a55;border-radius:8px;'
        'padding:.75rem 1.25rem;color:#4ade80;font-size:.82rem;margin-bottom:1.25rem">'
        '&#10003; All health checks passed — no missing summaries, full vector coverage.</div>'
        if health_ok else
        '<div style="background:#7c1d1d33;border:1px solid #ef444455;border-radius:8px;'
        'padding:.75rem 1.25rem;color:#f87171;font-size:.82rem;margin-bottom:1.25rem">'
        '&#9888; Health issues detected — see Health Alerts below.</div>'
    )

    # Extension rows
    by_ext = data.get("by_extension", [])
    max_ext = max((e["count"] for e in by_ext), default=1)
    ext_rows = "".join(
        f'<tr>'
        f'<td style="font-family:monospace;color:#bfdbfe">.{e["ext"]}</td>'
        f'<td style="text-align:right;font-weight:700;color:#e2e8f0">{e["count"]}</td>'
        f'<td style="text-align:right;color:#e2e8f0">{fmt_bytes(e["size_bytes"])}</td>'
        f'<td style="width:180px"><div style="background:#0f1117;border-radius:6px;height:8px;overflow:hidden">'
        f'<div style="width:{e["count"]/max_ext*100:.1f}%;height:100%;border-radius:6px;background:linear-gradient(90deg,#1d4ed8,#60a5fa)"></div>'
        f'</div></td></tr>'
        for e in by_ext
    ) or '<tr><td colspan="4" style="color:#94a3b8">No data</td></tr>'

    # Timeline bars
    by_day = data.get("indexed_by_day", [])
    max_day = max((d["count"] for d in by_day), default=1)
    day_bars = "".join(
        f'<div style="display:flex;flex-direction:column;align-items:center;gap:.2rem;flex:1;min-width:30px;max-width:60px">'
        f'<span style="font-size:.65rem;color:#cbd5e1">{d["count"]}</span>'
        f'<div style="height:{max(4,int(d["count"]/max_day*64))}px;width:100%;'
        f'background:linear-gradient(180deg,#60a5fa,#1d4ed8);border-radius:3px 3px 0 0"></div>'
        f'<span style="font-size:.6rem;color:#94a3b8;writing-mode:vertical-lr;transform:rotate(180deg)">{d["day"][-5:]}</span>'
        f'</div>'
        for d in by_day
    )
    timeline_html = (
        f'<div style="display:flex;align-items:flex-end;gap:3px;height:110px;padding-bottom:.5rem;overflow-x:auto">{day_bars}</div>'
        if by_day else '<p style="color:#94a3b8;font-size:.82rem">No indexing history available.</p>'
    )

    # Top files by chunks
    top_chunks = data.get("top_files_by_chunks", [])
    max_tc = max((f["chunk_count"] for f in top_chunks), default=1)
    top_chunk_rows = "".join(
        f'<tr>'
        f'<td style="max-width:420px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#f1f5f9">{f["filename"]}</td>'
        f'<td style="text-align:right;font-weight:700;color:#a78bfa">{f["chunk_count"]}</td>'
        f'<td style="width:140px"><div style="background:#0f1117;border-radius:6px;height:8px;overflow:hidden">'
        f'<div style="width:{f["chunk_count"]/max_tc*100:.1f}%;height:100%;border-radius:6px;background:linear-gradient(90deg,#7c3aed,#a78bfa)"></div>'
        f'</div></td></tr>'
        for f in top_chunks
    ) or '<tr><td colspan="3" style="color:#94a3b8">No data</td></tr>'

    # Recent files
    recent_rows = "".join(
        f'<tr>'
        f'<td style="max-width:380px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#f1f5f9">{f["filename"]}</td>'
        f'<td style="font-family:monospace;color:#bfdbfe">.{f["extension"]}</td>'
        f'<td style="color:#e2e8f0">{f["size_text"]}</td>'
        f'<td style="color:#cbd5e1;font-size:.75rem">{f["insert_dt"]}</td>'
        f'</tr>'
        for f in data.get("recent_files", [])
    ) or '<tr><td colspan="4" style="color:#94a3b8">No data</td></tr>'

    # Version rows
    ver_rows = "".join(
        f'<tr>'
        f'<td style="max-width:420px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#f1f5f9">{f["filename"]}</td>'
        f'<td style="text-align:right;font-weight:700;color:#34d399">{f["version_count"]}</td>'
        f'</tr>'
        for f in data.get("top_files_by_versions", [])
    ) or '<tr><td colspan="2" style="color:#94a3b8">No version history recorded.</td></tr>'

    total_ver   = data.get("total_version_records", 0)
    files_w_ver = data.get("files_with_versions", 0)
    avg_ver     = data.get("avg_versions_per_file", 0)

    # Contributor rows
    def _contrib(items):
        return "".join(
            f'<tr><td style="color:#f1f5f9">{i["name"]}</td>'
            f'<td style="text-align:right;font-weight:700;color:#fbbf24">{i["count"]}</td></tr>'
            for i in items
        ) or '<tr><td colspan="2" style="color:#94a3b8">No data</td></tr>'

    mod_rows = _contrib(data.get("top_modifiers", []))
    cre_rows = _contrib(data.get("top_creators",  []))

    # Table stats section
    table_stats = data.get("table_stats", {})
    def _src_badge(by_source):
        if not by_source:
            return "—"
        return " &nbsp; ".join(
            f'<span style="background:#1e293b;border:1px solid #334155;border-radius:4px;'
            f'padding:.1rem .45rem;font-size:.7rem;color:#94a3b8">'
            f'{src}: <b style="color:#e2e8f0">{cnt}</b></span>'
            for src, cnt in sorted(by_source.items())
        )

    def _tbl_total(info):
        t = info.get("total", "error")
        return "error" if t == "error" else f"{t:,}"

    tbl_rows = "".join(
        f'<tr>'
        f'<td style="font-family:monospace;font-size:.78rem;color:#bfdbfe">{tbl}</td>'
        f'<td style="text-align:right;font-weight:700;color:#e2e8f0">{_tbl_total(info)}</td>'
        f'<td>{_src_badge(info.get("by_source", {}))}</td>'
        f'</tr>'
        for tbl, info in table_stats.items()
    ) or '<tr><td colspan="3" style="color:#94a3b8">No data</td></tr>'

    # S3 section
    s3_total_files  = data.get("s3_total_files", 0)
    s3_total_size   = fmt_bytes(data.get("s3_total_size_bytes", 0))
    s3_by_ext       = data.get("s3_by_extension", [])
    s3_ext_count    = len(s3_by_ext)
    s3_max_ext      = max((e["count"] for e in s3_by_ext), default=1)
    s3_ext_rows     = "".join(
        f'<tr>'
        f'<td style="font-family:monospace;color:#fed7aa">.{e["ext"]}</td>'
        f'<td style="text-align:right;font-weight:700;color:#e2e8f0">{e["count"]}</td>'
        f'<td style="text-align:right;color:#e2e8f0">{fmt_bytes(e["size_bytes"])}</td>'
        f'<td style="width:180px"><div style="background:#0f1117;border-radius:6px;height:8px;overflow:hidden">'
        f'<div style="width:{e["count"]/s3_max_ext*100:.1f}%;height:100%;border-radius:6px;background:linear-gradient(90deg,#c2410c,#fb923c)"></div>'
        f'</div></td></tr>'
        for e in s3_by_ext
    ) or '<tr><td colspan="4" style="color:#94a3b8">No data</td></tr>'

    s3_by_day       = data.get("s3_indexed_by_day", [])
    s3_max_day      = max((d["count"] for d in s3_by_day), default=1)
    s3_day_bars     = "".join(
        f'<div style="display:flex;flex-direction:column;align-items:center;gap:.2rem;flex:1;min-width:30px;max-width:60px">'
        f'<span style="font-size:.65rem;color:#cbd5e1">{d["count"]}</span>'
        f'<div style="height:{max(4,int(d["count"]/s3_max_day*64))}px;width:100%;'
        f'background:linear-gradient(180deg,#fb923c,#c2410c);border-radius:3px 3px 0 0"></div>'
        f'<span style="font-size:.6rem;color:#94a3b8;writing-mode:vertical-lr;transform:rotate(180deg)">{d["day"][-5:]}</span>'
        f'</div>'
        for d in s3_by_day
    )
    s3_timeline_html = (
        f'<div style="display:flex;align-items:flex-end;gap:3px;height:110px;padding-bottom:.5rem;overflow-x:auto">{s3_day_bars}</div>'
        if s3_by_day else '<p style="color:#94a3b8;font-size:.82rem">No S3 indexing history available.</p>'
    )
    s3_recent_rows  = "".join(
        f'<tr>'
        f'<td style="max-width:320px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#f1f5f9">{f["filename"]}</td>'
        f'<td style="font-family:monospace;color:#fed7aa">.{f["extension"]}</td>'
        f'<td style="color:#e2e8f0">{f["size_text"]}</td>'
        f'<td style="color:#94a3b8;font-size:.75rem">{f["storage_class"]}</td>'
        f'<td style="color:#cbd5e1;font-size:.75rem">{f["insert_dt"]}</td>'
        f'</tr>'
        for f in data.get("s3_recent_files", [])
    ) or '<tr><td colspan="5" style="color:#94a3b8">No data</td></tr>'

    # Shared breakdown
    shared_map = data.get("shared_breakdown", {})
    shared_rows = "".join(
        f'<tr><td style="color:#f1f5f9">{k}</td>'
        f'<td style="text-align:right;font-weight:700;color:#f1f5f9">{v}</td></tr>'
        for k, v in shared_map.items()
    ) or '<tr><td colspan="2" style="color:#94a3b8">No data</td></tr>'

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>DB Analysis &mdash; {ts}</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',system-ui,sans-serif;background:#0f1117;color:#e2e8f0;padding:2rem;min-height:100vh}}
h1{{font-size:1.5rem;font-weight:700;color:#f8fafc;margin-bottom:.25rem}}
h2{{font-size:1rem;color:#e2e8f0;font-weight:400;margin-bottom:.75rem}}
.meta{{font-size:.75rem;color:#cbd5e1;margin-top:.4rem;margin-bottom:1.5rem;padding-bottom:1rem;border-bottom:1px solid #1e293b}}
.grid-2{{display:grid;grid-template-columns:repeat(2,1fr);gap:1rem;margin-bottom:1.5rem}}
.grid-3{{display:grid;grid-template-columns:repeat(3,1fr);gap:1rem;margin-bottom:1rem}}
.grid-4{{display:grid;grid-template-columns:repeat(4,1fr);gap:1rem;margin-bottom:1.5rem}}
.card{{background:#1e293b;border:1px solid #334155;border-radius:10px;padding:1.25rem}}
.card-green{{border-color:#22c55e44}}.card-blue{{border-color:#3b82f644}}
.card-purple{{border-color:#a855f744}}.card-red{{border-color:#ef444444}}
.card-teal{{border-color:#2dd4bf44}}.card-orange{{border-color:#f9731644}}
.kpi{{display:flex;flex-direction:column;gap:.2rem}}
.kpi .lbl{{font-size:.72rem;color:#e2e8f0;text-transform:uppercase;letter-spacing:.06em}}
.kpi .val{{font-size:2rem;font-weight:700;line-height:1}}
.kpi .sub{{font-size:.78rem;margin-top:.3rem;color:#cbd5e1}}
.green{{color:#22c55e}}.blue{{color:#60a5fa}}.purple{{color:#a78bfa}}
.teal{{color:#2dd4bf}}.gray{{color:#e2e8f0}}
.section-title{{font-size:.72rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#94a3b8;display:flex;align-items:center;gap:.5rem;margin-bottom:.75rem}}
.section-title::after{{content:'';flex:1;height:1px;background:#1e293b}}
section{{margin-bottom:1.75rem}}
table{{width:100%;border-collapse:collapse;font-size:.82rem}}
th{{padding:.55rem .8rem;text-align:left;color:#cbd5e1;font-size:.72rem;font-weight:600;text-transform:uppercase;letter-spacing:.06em;border-bottom:1px solid #334155;background:#0f1117}}
td{{padding:.55rem .8rem;border-bottom:1px solid #1e293b;color:#f1f5f9;vertical-align:middle}}
tr:last-child td{{border-bottom:none}}
tr:hover td{{background:#1e293b55}}
</style>
</head>
<body>
<h1>Database Analysis</h1>
<h2>HANA Cloud &middot; All Tables Snapshot</h2>
<div class="meta">Generated: <code>{ts}</code> &nbsp;|&nbsp; Tables: BRAIN_SP_FILES &middot; BRAIN_TEXT_CHUNKS &middot; BRAIN_VECTORS &middot; BRAIN_SP_FILES_CHUNK &middot; BRAIN_SP_FILES_VERSIONS</div>

{health_banner}

<section>
<div class="section-title">Overview</div>
<div class="grid-4">
  <div class="card card-green">
    <div class="kpi"><span class="lbl">Files Indexed</span>
    <span class="val green">{total_files:,}</span>
    <span class="sub">{fmt_bytes(total_bytes)} total raw size</span></div>
  </div>
  <div class="card card-blue">
    <div class="kpi"><span class="lbl">Text Chunks</span>
    <span class="val blue">{total_chunks:,}</span>
    <span class="sub">avg {avg_chunks_per_file} per file</span></div>
  </div>
  <div class="card card-purple">
    <div class="kpi"><span class="lbl">Vectors</span>
    <span class="val purple">{total_vectors:,}</span>
    <span class="sub">{vector_coverage}% chunk coverage</span>
    <div style="background:#0f1117;border-radius:6px;height:8px;overflow:hidden;margin-top:.4rem">
    <div style="width:{vector_coverage}%;height:100%;border-radius:6px;background:linear-gradient(90deg,#7c3aed,#a78bfa)"></div></div>
    </div>
  </div>
  <div class="card card-teal">
    <div class="kpi"><span class="lbl">Avg Chunk Length</span>
    <span class="val teal">{avg_chunk_chars:,}</span>
    <span class="sub">characters per chunk</span></div>
  </div>
</div>
</section>

<section>
<div class="section-title">Health Alerts</div>
<div class="grid-2">
  <div class="card {cs_card}">
    <div class="kpi"><span class="lbl">Files Missing Chunk Summary</span>
    <span class="val {cs_col}">{missing_cs:,}</span>
    <span class="sub">{cs_sub}</span></div>
  </div>
  <div class="card {vec_card}">
    <div class="kpi"><span class="lbl">Chunks Without Vectors</span>
    <span class="val {vec_col}">{chunks_no_v:,}</span>
    <span class="sub">{vec_sub}</span></div>
  </div>
</div>
</section>

<section>
<div class="section-title">Table Stats</div>
<div class="card">
<table>
<thead><tr><th>Table</th><th style="text-align:right">Row Count</th><th>By Source</th></tr></thead>
<tbody>{tbl_rows}</tbody>
</table>
</div>
</section>

<section>
<div class="section-title">File Types Breakdown</div>
<div class="card">
<table>
<thead><tr><th>Extension</th><th style="text-align:right">Files</th><th style="text-align:right">Total Size</th><th>Distribution</th></tr></thead>
<tbody>{ext_rows}</tbody>
</table>
</div>
</section>

<section>
<div class="section-title">Shared vs Private</div>
<div class="card" style="max-width:360px">
<table>
<thead><tr><th>Shared Value</th><th style="text-align:right">Files</th></tr></thead>
<tbody>{shared_rows}</tbody>
</table>
</div>
</section>

<section>
<div class="section-title">Indexing Timeline (files per day)</div>
<div class="card">{timeline_html}</div>
</section>

<section>
<div class="section-title">Top 10 Files by Chunk Count</div>
<div class="card">
<table>
<thead><tr><th>Filename</th><th style="text-align:right">Chunks</th><th>Distribution</th></tr></thead>
<tbody>{top_chunk_rows}</tbody>
</table>
</div>
</section>

<section>
<div class="section-title">Recently Indexed Files (Last 10)</div>
<div class="card">
<table>
<thead><tr><th>Filename</th><th>Type</th><th>Size</th><th>Indexed At (UTC)</th></tr></thead>
<tbody>{recent_rows}</tbody>
</table>
</div>
</section>

<section>
<div class="section-title">Version History</div>
<div class="grid-3">
  <div class="card card-teal">
    <div class="kpi"><span class="lbl">Total Version Records</span>
    <span class="val teal">{total_ver:,}</span></div>
  </div>
  <div class="card card-blue">
    <div class="kpi"><span class="lbl">Files With Versions</span>
    <span class="val blue">{files_w_ver:,}</span></div>
  </div>
  <div class="card card-purple">
    <div class="kpi"><span class="lbl">Avg Versions / File</span>
    <span class="val purple">{avg_ver}</span></div>
  </div>
</div>
<div class="card">
<table>
<thead><tr><th>Filename</th><th style="text-align:right">Version Records</th></tr></thead>
<tbody>{ver_rows}</tbody>
</table>
</div>
</section>

<section>
<div class="section-title">Contributors</div>
<div class="grid-2">
  <div class="card">
    <div style="font-size:.72rem;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#cbd5e1;margin-bottom:.75rem">Top Modifiers</div>
    <table><thead><tr><th>Name</th><th style="text-align:right">Files</th></tr></thead>
    <tbody>{mod_rows}</tbody></table>
  </div>
  <div class="card">
    <div style="font-size:.72rem;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#cbd5e1;margin-bottom:.75rem">Top Creators</div>
    <table><thead><tr><th>Name</th><th style="text-align:right">Files</th></tr></thead>
    <tbody>{cre_rows}</tbody></table>
  </div>
</div>
</section>

<section>
<div class="section-title">S3 Object Store</div>
<div class="grid-2">
  <div class="card card-orange">
    <div class="kpi"><span class="lbl">S3 Files Indexed</span>
    <span class="val" style="color:#fb923c">{s3_total_files:,}</span>
    <span class="sub">{s3_total_size}</span></div>
  </div>
  <div class="card card-orange">
    <div class="kpi"><span class="lbl">S3 File Types</span>
    <span class="val" style="color:#fb923c">{s3_ext_count}</span>
    <span class="sub">distinct extensions</span></div>
  </div>
</div>
<div class="card" style="margin-bottom:1rem">
<table>
<thead><tr><th>Extension</th><th style="text-align:right">Files</th><th style="text-align:right">Total Size</th><th>Distribution</th></tr></thead>
<tbody>{s3_ext_rows}</tbody>
</table>
</div>
<div class="card" style="margin-bottom:1rem">
<div style="font-size:.72rem;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#94a3b8;margin-bottom:.75rem">S3 Indexing Timeline</div>
{s3_timeline_html}
</div>
<div class="card">
<div style="font-size:.72rem;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#94a3b8;margin-bottom:.75rem">Recently Indexed S3 Files (Last 10)</div>
<table>
<thead><tr><th>Filename</th><th>Type</th><th>Size</th><th>Storage Class</th><th>Indexed At (UTC)</th></tr></thead>
<tbody>{s3_recent_rows}</tbody>
</table>
</div>
</section>

</body>
</html>"""


def _generate_analysis_html(run_summary: dict, save_dir: str) -> str:
    ts         = run_summary["ts"]
    total      = run_summary["total_items"]
    vectorized = run_summary["vectorized_count"]
    n_chunks   = run_summary["chunk_count"]
    skips      = run_summary["skips"]
    retries    = run_summary["retries"]
    file_sizes = run_summary["file_sizes"]
    duration_s = run_summary.get("duration_s", 0)
    log_path   = run_summary.get("log_path")
    db_stats   = run_summary.get("db_stats", {})

    no_text  = [s for s in skips if s["outcome"] == "no_text"]
    timeouts = [s for s in skips if s["outcome"] == "timeout_exhausted"]
    http_err = [s for s in skips if s["outcome"].startswith("http_")]
    other    = [s for s in skips if s["outcome"] in ("no_url", "error")]
    coverage = (vectorized / total * 100) if total else 0
    mins     = int(duration_s // 60)
    secs     = int(duration_s % 60)
    dur_str  = f"{mins}m {secs}s" if mins else f"{secs}s"

    def fmt_bytes(b):
        if not b: return "—"
        for u in ("B", "KB", "MB", "GB"):
            if b < 1024: return f"{b:.1f} {u}"
            b /= 1024
        return f"{b:.1f} TB"

    def skip_rows(items, color):
        return "".join(
            f'<tr><td>{s["filename"]}</td><td>{s["size"]}</td>'
            f'<td><span class="badge" style="background:{color}22;color:{color};border:1px solid {color}44">'
            f'{s["outcome"].replace("_"," ").upper()}</span></td></tr>'
            for s in items
        )

    retry_rows = "".join(
        f'<tr><td>{r["filename"]}</td><td>{r["size"]}</td>'
        f'<td style="color:{"#f97316" if r["attempts"]==3 else "#d8b4fe"};font-weight:700">{r["attempts"]}/3</td>'
        f'<td><span class="badge badge-green">recovered</span></td></tr>'
        for r in retries
    )

    skip_section = (
        (f'<h3 style="margin-top:1rem;margin-bottom:.5rem">Timeout Exhausted ({len(timeouts)})</h3>'
         f'<table><thead><tr><th>File</th><th>Size</th><th>Status</th></tr></thead>'
         f'<tbody>{skip_rows(timeouts,"#ef4444")}</tbody></table>' if timeouts else "") +
        (f'<h3 style="margin-top:1rem;margin-bottom:.5rem">HTTP Errors ({len(http_err)})</h3>'
         f'<table><thead><tr><th>File</th><th>Size</th><th>Status</th></tr></thead>'
         f'<tbody>{skip_rows(http_err,"#f97316")}</tbody></table>' if http_err else "") +
        (f'<h3 style="margin-top:1rem;margin-bottom:.5rem">No Text Extracted ({len(no_text)})</h3>'
         f'<table><thead><tr><th>File</th><th>Size</th><th>Status</th></tr></thead>'
         f'<tbody>{skip_rows(no_text,"#60a5fa")}</tbody></table>' if no_text else "") +
        (f'<h3 style="margin-top:1rem;margin-bottom:.5rem">Other ({len(other)})</h3>'
         f'<table><thead><tr><th>File</th><th>Size</th><th>Status</th></tr></thead>'
         f'<tbody>{skip_rows(other,"#94a3b8")}</tbody></table>' if other else "")
    )

    log_row = (
        f'<tr><td style="font-family:monospace;font-size:.78rem;color:#94a3b8">{os.path.basename(log_path)}</td>'
        f'<td style="font-weight:700;color:#e2e8f0">{fmt_bytes(os.path.getsize(log_path)) if os.path.exists(log_path) else "—"}</td></tr>'
        if log_path else ""
    )
    file_size_rows = "".join(
        f'<tr><td style="font-family:monospace;font-size:.78rem;color:#94a3b8">{n}</td>'
        f'<td style="font-weight:700;color:#e2e8f0">{fmt_bytes(s)}</td></tr>'
        for n, s in file_sizes.items()
    ) + log_row

    retry_section = (
        f'<section><div class="section-title">Retry Detail ({len(retries)} files needed &gt;1 attempt)</div>'
        f'<div class="card card-orange"><table><thead><tr><th>File</th><th>Size</th>'
        f'<th>Attempts Used</th><th>Outcome</th></tr></thead>'
        f'<tbody>{retry_rows}</tbody></table></div></section>'
        if retries else ""
    )

    _db_label_map = {
        "files":          "BRAIN_SP_FILES",
        "chunks":         "BRAIN_TEXT_CHUNKS",
        "vectors":        "BRAIN_VECTORS",
        # "file_registry":  "BRAIN_FILE_REGISTRY",  # disabled
    }
    def _fmt_count(v):
        return f"{v:,}" if isinstance(v, int) else str(v)

    if db_stats:
        _db_rows = ""
        for key, label in _db_label_map.items():
            entry = db_stats.get(key)
            if entry is None:
                continue
            total_count = entry["total"]
            by_src      = entry.get("by_source", {})
            sp_count    = by_src.get("sharepoint", "—") if by_src else "—"
            s3_count    = by_src.get("s3",         "—") if by_src else "—"
            _db_rows += (
                f'<tr>'
                f'<td style="font-family:monospace;font-size:.78rem;color:#94a3b8">{label}</td>'
                f'<td style="font-weight:700;color:#e2e8f0;text-align:right">{_fmt_count(total_count)}</td>'
                f'<td style="color:#60a5fa;text-align:right">{_fmt_count(sp_count)}</td>'
                f'<td style="color:#34d399;text-align:right">{_fmt_count(s3_count)}</td>'
                f'</tr>'
            )
        db_section = (
            f'<section><div class="section-title">Database Tables</div>'
            f'<div class="card" style="max-width:720px">'
            f'<table><thead><tr><th>Table</th>'
            f'<th style="text-align:right">Total</th>'
            f'<th style="text-align:right;color:#60a5fa">SharePoint</th>'
            f'<th style="text-align:right;color:#34d399">S3</th></tr></thead>'
            f'<tbody>{_db_rows}</tbody></table></div></section>'
        )
    else:
        db_section = (
            '<section><div class="section-title">Database Tables</div>'
            '<div class="card"><p style="color:#64748b;font-size:.82rem;padding:.5rem 0">'
            'DB stats unavailable — could not connect to HANA at report time.</p></div></section>'
        )

    skip_color    = "card-orange" if (timeouts or http_err or other) else "card-green"
    retry_color   = "card-orange" if retries else "card-green"
    skip_val_col  = "orange" if (timeouts or http_err or other) else "green"
    retry_val_col = "orange" if retries else "green"

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Vectorize Analysis — {ts}</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',system-ui,sans-serif;background:#0f1117;color:#e2e8f0;padding:2rem;min-height:100vh}}
h1{{font-size:1.5rem;font-weight:700;color:#f8fafc;margin-bottom:.25rem}}
h2{{font-size:1rem;color:#94a3b8;font-weight:400;margin-bottom:.75rem}}
h3{{font-size:.85rem;font-weight:600;color:#cbd5e1;text-transform:uppercase;letter-spacing:.06em}}
.meta{{font-size:.75rem;color:#64748b;margin-top:.4rem;margin-bottom:1.75rem;padding-bottom:1rem;border-bottom:1px solid #1e293b}}
.grid-4{{display:grid;grid-template-columns:repeat(4,1fr);gap:1rem;margin-bottom:1.5rem}}
.card{{background:#1e293b;border:1px solid #334155;border-radius:10px;padding:1.25rem}}
.card-green{{border-color:#22c55e44}}.card-blue{{border-color:#3b82f644}}
.card-purple{{border-color:#a855f744}}.card-orange{{border-color:#f9731644}}
.kpi{{display:flex;flex-direction:column;gap:.2rem}}
.kpi .lbl{{font-size:.72rem;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em}}
.kpi .val{{font-size:2rem;font-weight:700;line-height:1}}
.kpi .sub{{font-size:.78rem;margin-top:.2rem}}
.green{{color:#22c55e}}.blue{{color:#60a5fa}}.orange{{color:#fb923c}}.gray{{color:#94a3b8}}
.bar-wrap{{background:#0f1117;border-radius:6px;height:8px;overflow:hidden;margin-top:.4rem}}
.bar{{height:100%;border-radius:6px}}
.section-title{{font-size:.72rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#475569;display:flex;align-items:center;gap:.5rem;margin-bottom:.75rem}}
.section-title::after{{content:'';flex:1;height:1px;background:#1e293b}}
section{{margin-bottom:1.75rem}}
table{{width:100%;border-collapse:collapse;font-size:.82rem}}
th{{padding:.55rem .8rem;text-align:left;color:#64748b;font-size:.72rem;font-weight:600;text-transform:uppercase;letter-spacing:.06em;border-bottom:1px solid #334155;background:#0f1117}}
td{{padding:.55rem .8rem;border-bottom:1px solid #1e293b;color:#cbd5e1;vertical-align:middle}}
tr:last-child td{{border-bottom:none}}
tr:hover td{{background:#1e293b55}}
.badge{{display:inline-block;padding:.12rem .5rem;border-radius:999px;font-size:.7rem;font-weight:600}}
.badge-green{{background:#14532d55;color:#4ade80;border:1px solid #16a34a55}}
</style>
</head>
<body>
<h1>Vectorize Analysis</h1>
<h2>SharePoint RAG API · Auto-generated run report</h2>
<div class="meta">
  Run timestamp: <code>{ts}</code> &nbsp;|&nbsp; Duration: <strong>{dur_str}</strong> &nbsp;|&nbsp; Files scanned: {total}
</div>

<section>
<div class="section-title">Summary</div>
<div class="grid-4">
  <div class="card card-green">
    <div class="kpi">
      <span class="lbl">Files Vectorized</span>
      <span class="val green">{vectorized}</span>
      <span class="sub gray">{coverage:.1f}% of {total}</span>
      <div class="bar-wrap"><div class="bar" style="width:{coverage:.1f}%;background:linear-gradient(90deg,#16a34a,#22c55e)"></div></div>
    </div>
  </div>
  <div class="card card-blue">
    <div class="kpi">
      <span class="lbl">Chunk Records</span>
      <span class="val blue">{n_chunks:,}</span>
      <span class="sub gray">semantic chunks embedded</span>
    </div>
  </div>
  <div class="card {skip_color}">
    <div class="kpi">
      <span class="lbl">Total Skipped</span>
      <span class="val {skip_val_col}">{len(skips)}</span>
      <span class="sub gray">{len(no_text)} no-text · {len(timeouts)} timeout · {len(http_err)} HTTP err</span>
    </div>
  </div>
  <div class="card {retry_color}">
    <div class="kpi">
      <span class="lbl">Retries Triggered</span>
      <span class="val {retry_val_col}">{len(retries)}</span>
      <span class="sub gray">{"files >1 attempt — all recovered" if retries else "all succeeded on attempt 1"}</span>
    </div>
  </div>
</div>
</section>

{retry_section}

<section>
<div class="section-title">Skipped Files ({len(skips)} total)</div>
<div class="card">
{skip_section or '<p style="color:#64748b;font-size:.82rem;padding:.5rem 0">No files skipped — all files downloaded and extracted successfully.</p>'}
</div>
</section>

{db_section}

<section>
<div class="section-title">Output Files</div>
<div class="card" style="max-width:640px">
<table><thead><tr><th>File</th><th>Size</th></tr></thead>
<tbody>{file_size_rows}</tbody></table>
</div>
</section>

<div class="meta" style="margin-top:2rem;text-align:center">
  Auto-generated by pytesting/catalyst/for_SAP/api.py &nbsp;·&nbsp; {ts}
</div>
</body>
</html>"""

    path = os.path.join(save_dir, f"sp_analysis_{ts}.html")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)
    log.info(f"Analysis HTML written to {path}")
    return path


# ── Chunk JSONL reader ────────────────────────────────────────────────────────

def _read_chunks_from_jsonl(path: str, filename: str) -> list:
    log.info(f"Reading chunks from {path}")
    chunks = []
    try:
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    record = json.loads(line)
                    if record.get("filename") == filename:
                        chunks.append(record)
                except json.JSONDecodeError:
                    continue
    except FileNotFoundError:
        pass

    log.info(f"Returning {len(chunks)} chunks for {filename}")
    return chunks


# ── File downloader (bytes + base64 chunks) ───────────────────────────────────

def _download_file_content(item: dict, save_dir: str = None) -> bytes:
    url       = item.get("BaseURL")
    file_name = item.get("name")

    if not url:
        log.warning(f"No BaseURL for {file_name}")
        return b""

    if save_dir:
        os.makedirs(save_dir, exist_ok=True)
        json_path = os.path.join(save_dir, "files_list_chunk.json")
        if not os.path.exists(json_path):
            with open(json_path, "w"):
                pass
        log.info(f"Saving chunks to {json_path}")

    print(f"Downloading: {file_name} ({item.get('size', '—')})")
    log.info(f"Downloading: {file_name} ({item.get('size', '—')})")

    try:
        chunks = []
        with requests.get(url, headers=_headers(), stream=True, timeout=(15, 300)) as r:
            r.raise_for_status()
            for i, chunk in enumerate(r.iter_content(chunk_size=8192), start=1):
                if chunk:
                    chunks.append(chunk)
                    encoded    = base64.b64encode(chunk).decode("utf-8")
                    chunk_list = [{
                        "item_id":        item.get("id"),
                        "filename":       item.get("name"),
                        "chunk_datetime": datetime.now().isoformat(),
                        "chunk_size":     len(chunk),
                        "chunk_id":       i,
                        "chunk_data":     encoded,
                    }]
                    if save_dir:
                        _write_to_json(chunk_list, path=os.path.join(save_dir, "files_list_chunk.json"))
                    print(f"  chunk [{i}]: {len(chunk)} bytes | {encoded[:20]}...")

        total_chunks = len(chunks)
        content      = b"".join(chunks)
        print(f"  Downloaded {len(content) / 1024:.1f} KB  ({total_chunks} chunks)")
        log.info(f"  Downloaded {len(content) / 1024:.1f} KB  ({total_chunks} chunks)")

        if save_dir:
            save_path = os.path.join(save_dir, file_name)
            with open(save_path, "wb") as f:
                f.write(content)
            print(f"  Saved to: {save_path}")
            log.info(f"  Saved to: {save_path}")

        return content

    except requests.exceptions.ReadTimeout:
        print(f"[TIMEOUT] {file_name} — skipping")
        log.warning(f"[TIMEOUT] {file_name} — skipping")
    except requests.exceptions.HTTPError as e:
        sc = e.response.status_code if e.response is not None else "?"
        print(f"[HTTP ERROR] {file_name} — {sc}")
        log.error(f"[HTTP ERROR] {file_name} — {sc}")
    except Exception as e:
        print(f"[ERROR] {file_name} — {e}")
        log.error(f"[ERROR] {file_name} — {e}")
    return b""


# ── Chunk rebuilder ───────────────────────────────────────────────────────────

def _rebuild_file(chunks: list, save_dir: str = None) -> bytes:
    if not chunks:
        print("[ERROR] No chunks provided")
        log.error("No chunks provided")
        return b""

    try:
        sorted_chunks = sorted(chunks, key=lambda x: x.get("chunk_id", 0))
        file_name     = sorted_chunks[0].get("filename")
        total         = len(sorted_chunks)

        print(f"Rebuilding: {file_name} ({total} chunks)")
        log.info(f"Rebuilding: {file_name} ({total} chunks)")

        decoded = []
        for c in sorted_chunks:
            chunk_id = c.get("chunk_id")
            original = base64.b64decode(c.get("chunk_data", "").encode("utf-8"))
            decoded.append(original)
            print(f"  chunk [{chunk_id}]: {len(original)} bytes decoded")
            log.info(f"  chunk [{chunk_id}]: {len(original)} bytes decoded")

        content = b"".join(decoded)
        print(f"  Rebuilt {len(content) / 1024:.1f} KB from {total} chunks")
        log.info(f"  Rebuilt {len(content) / 1024:.1f} KB from {total} chunks")

        if save_dir:
            os.makedirs(save_dir, exist_ok=True)
            save_path = os.path.join(save_dir, file_name)
            with open(save_path, "wb") as f:
                f.write(content)
            print(f"  Saved to: {save_path}")
            log.info(f"  Saved to: {save_path}")

        return content

    except Exception as e:
        print(f"[ERROR] Rebuild failed — {e}")
        log.error(f"[ERROR] Rebuild failed — {e}")
        return b""


# ── RAG: text extraction, chunking, vectorization ────────────────────────────

def _extract_text(content: bytes, filename: str) -> str:
    log.info(f"Extracting text from {filename}")
    ext = filename.rsplit(".", 1)[-1].lower()
    try:
        if ext in ("docx", "dotx"):
            log.info("Extracting docx/dotx.")
            doc = Document(io.BytesIO(content))
            return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

        elif ext == "doc":
            log.info("Extracting doc (binary Word).")
            import olefile, re
            with olefile.OleFileIO(io.BytesIO(content)) as ole:
                if ole.exists("WordDocument"):
                    raw  = ole.openstream("WordDocument").read()
                    text = raw.decode("utf-16-le", errors="ignore")
                    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
                    text = re.sub(r" {3,}", " ", text).strip()
                    return text
            return ""

        elif ext in ("xlsx", "xlsm"):
            log.info("Extracting xlsx.")
            try:
                wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True)
            except Exception as e:
                if "zip" in str(e).lower():
                    log.warning("  xlsx is not a zip file — retrying with xlrd (likely legacy XLS binary)")
                    import xlrd
                    wb = xlrd.open_workbook(file_contents=content)
                    lines = []
                    for sheet in wb.sheets():
                        for row_idx in range(sheet.nrows):
                            line = " | ".join(
                                str(sheet.cell_value(row_idx, col))
                                for col in range(sheet.ncols)
                                if sheet.cell_value(row_idx, col) not in (None, "")
                            )
                            if line.strip():
                                lines.append(line)
                    return "\n".join(lines)
                raise
            lines = []
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    line = " | ".join(str(c) for c in row if c is not None)
                    if line.strip():
                        lines.append(line)
            return "\n".join(lines)

        elif ext == "xls":
            log.info("Extracting xls (legacy binary).")
            import xlrd
            wb    = xlrd.open_workbook(file_contents=content)
            lines = []
            for sheet in wb.sheets():
                for row_idx in range(sheet.nrows):
                    line = " | ".join(
                        str(sheet.cell_value(row_idx, col))
                        for col in range(sheet.ncols)
                        if sheet.cell_value(row_idx, col) not in (None, "")
                    )
                    if line.strip():
                        lines.append(line)
            return "\n".join(lines)

        elif ext == "pdf":
            log.info("Extracting pdf.")
            text = []
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for page in pdf.pages:
                    text.append(page.extract_text() or "")
            return "\n".join(text)

        elif ext in ("pptx", "pptm"):
            log.info("Extracting pptx/pptm.")
            prs   = Presentation(io.BytesIO(content))
            lines = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            t = para.text.strip()
                            if t:
                                lines.append(t)
            return "\n".join(lines)

        elif ext == "ppt":
            log.info("Extracting ppt (binary PowerPoint).")
            import olefile, re
            texts = []
            with olefile.OleFileIO(io.BytesIO(content)) as ole:
                for entry in ole.listdir():
                    try:
                        raw     = ole.openstream(entry).read()
                        decoded = raw.decode("utf-16-le", errors="ignore")
                        cleaned = re.sub(r"[\x00-\x1f]", " ", decoded).strip()
                        if len(cleaned) > 20:
                            texts.append(cleaned)
                    except Exception:
                        continue
            return "\n".join(texts)

        elif ext == "rtf":
            log.info("Extracting rtf.")
            from striprtf.striprtf import rtf_to_text
            return rtf_to_text(content.decode("latin-1", errors="replace"))

        elif ext == "odt":
            log.info("Extracting odt.")
            import zipfile, xml.etree.ElementTree as ET
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                xml_bytes = zf.read("content.xml")
            root  = ET.fromstring(xml_bytes)
            texts = [t.strip() for t in root.itertext() if t.strip()]
            return "\n".join(texts)

        elif ext == "eml":
            log.info("Extracting eml (email).")
            import email as _email
            from email import policy as _policy
            msg  = _email.message_from_bytes(content, policy=_policy.default)
            subj = msg.get("subject", "")
            frm  = msg.get("from", "")
            to   = msg.get("to", "")
            cc   = msg.get("cc", "")
            date = msg.get("date", "")
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    ct = part.get_content_type()
                    cd = str(part.get("content-disposition", ""))
                    if "attachment" in cd:
                        continue
                    if ct == "text/plain":
                        body += part.get_payload(decode=True).decode(
                            part.get_content_charset() or "utf-8", errors="replace"
                        ) + "\n"
                    elif ct == "text/html" and not body:
                        raw_html = part.get_payload(decode=True).decode(
                            part.get_content_charset() or "utf-8", errors="replace"
                        )
                        body += _extract_text(raw_html.encode("utf-8"), "body.html") + "\n"
            else:
                body = msg.get_payload(decode=True).decode(
                    msg.get_content_charset() or "utf-8", errors="replace"
                )
            lines = [f"From: {frm}", f"To: {to}"]
            if cc:   lines.append(f"CC: {cc}")
            if date: lines.append(f"Date: {date}")
            lines.append(f"Subject: {subj}")
            lines.append("")
            lines.append(body.strip())
            return "\n".join(lines)

        elif ext == "msg":
            log.info("Extracting msg (Outlook).")
            import extract_msg as _emsg
            with _emsg.openMsg(io.BytesIO(content)) as msg:
                lines = [
                    f"From: {msg.sender or ''}",
                    f"To: {msg.to or ''}",
                ]
                if msg.cc:   lines.append(f"CC: {msg.cc}")
                if msg.date: lines.append(f"Date: {msg.date}")
                lines.append(f"Subject: {msg.subject or ''}")
                lines.append("")
                lines.append((msg.body or "").strip())
            return "\n".join(lines)

        elif ext in ("abap", "txt", "md", "py", "js", "ts", "sql", "yaml", "yml", "ini", "env", "toml", "log", "cfg", "java", "cs"):
            log.info(f"Extracting {ext} (plain-text).")
            return content.decode("utf-8", errors="replace")

        elif ext == "csv":
            log.info("Extracting csv.")
            import csv as _csv
            text_io = io.StringIO(content.decode("utf-8", errors="replace"))
            reader  = _csv.reader(text_io)
            lines   = [" | ".join(cell for cell in row if cell.strip()) for row in reader]
            return "\n".join(line for line in lines if line.strip())

        elif ext == "json":
            log.info("Extracting json.")
            try:
                parsed = json.loads(content.decode("utf-8", errors="replace"))
                return json.dumps(parsed, indent=2, ensure_ascii=False)
            except Exception:
                return content.decode("utf-8", errors="replace")

        elif ext == "xml":
            log.info("Extracting xml.")
            import xml.etree.ElementTree as ET
            try:
                root  = ET.fromstring(content.decode("utf-8", errors="replace"))
                texts = [t.strip() for t in root.itertext() if t.strip()]
                return "\n".join(texts)
            except ET.ParseError:
                return content.decode("utf-8", errors="replace")

        elif ext in ("html", "htm"):
            log.info("Extracting html.")
            from html.parser import HTMLParser
            class _StripTags(HTMLParser):
                def __init__(self):
                    super().__init__()
                    self._parts = []
                    self._skip  = False
                def handle_starttag(self, tag, attrs):
                    if tag in ("script", "style"):
                        self._skip = True
                def handle_endtag(self, tag):
                    if tag in ("script", "style"):
                        self._skip = False
                def handle_data(self, data):
                    if not self._skip and data.strip():
                        self._parts.append(data.strip())
            parser = _StripTags()
            parser.feed(content.decode("utf-8", errors="replace"))
            return "\n".join(parser._parts)

        elif ext == "zip":
            log.info("Extracting zip (processing supported files inside).")
            import zipfile
            texts = []
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                for entry in zf.infolist():
                    if entry.is_dir():
                        continue
                    inner_name = entry.filename
                    inner_ext  = inner_name.rsplit(".", 1)[-1].lower() if "." in inner_name else ""
                    if inner_ext == "zip":
                        log.info(f"  Skipping nested zip: {inner_name}")
                        continue
                    try:
                        inner_text = _extract_text(zf.read(entry), inner_name)
                        if inner_text.strip():
                            texts.append(f"=== {inner_name} ===\n{inner_text}")
                        else:
                            log.info(f"  No text from {inner_name} (unsupported or empty)")
                    except Exception as e:
                        log.warning(f"  Could not extract {inner_name} from zip: {e}")
            return "\n\n".join(texts)

    except Exception as e:
        print(f"[ERROR] Text extraction failed for {filename} — {e}")
        log.error(f"[ERROR] Text extraction failed for {filename} — {e}")
    return ""


def _semantic_chunks(text: str, max_tokens: int = 512) -> list:
    max_chars  = max_tokens * 4
    paragraphs = [p.strip() for p in text.split("\n") if p.strip()]
    chunks     = []
    current    = []
    length     = 0

    log.info(f"Chunking into ~{max_tokens} token segments")

    for para in paragraphs:
        if length + len(para) > max_chars and current:
            chunks.append("\n".join(current))
            current = []
            length  = 0
        current.append(para)
        length += len(para)

    if current:
        chunks.append("\n".join(current))

    log.info(f"Chunked into {len(chunks)} segments")

    return chunks


def _upload_run_artifacts(paths: list[str | None], logger: logging.Logger = None) -> dict:
    """
    Upload a list of local file paths to S3 under runs/YYYYMMDD/.
    Non-fatal — logs warnings on failure, never raises.
    Returns {filename: s3_key} for successfully uploaded files.
    """
    _log     = logger or log
    uploaded = {}
    day      = datetime.now().strftime("%Y%m%d")
    for path in paths:
        if not path or not os.path.exists(path):
            continue
        filename = os.path.basename(path)
        s3_key   = f"runs/{day}/{filename}"
        try:
            # Flush any open file handlers writing to this path before uploading
            for h in logging.root.handlers + (_log.handlers if _log else []):
                if isinstance(h, logging.FileHandler) and h.baseFilename == path:
                    h.flush()
            s3.upload_file(path, s3_key)
            uploaded[filename] = s3_key
            _log.info(f"  S3 upload ok — {filename} → {s3_key}")
        except Exception as e:
            _log.warning(f"  S3 upload failed for {filename}: {type(e).__name__}: {e}")
    return uploaded


def _filter_stale_items(items: list[dict], logger: logging.Logger = None) -> tuple[list[dict], int]:
    """
    Compare SharePoint items against DB. Return (items_to_process, skipped_count).
    An item is skipped only when it already exists in the DB with the same last_modified timestamp.
    Falls back to processing all items if the DB map cannot be fetched.
    """
    _log = logger or log
    try:
        db_map = db.get_modified_map()
    except Exception as e:
        _log.warning(f"Could not fetch DB modified map — processing all items: {e}")
        return items, 0

    def _norm_ts(ts: str) -> str:
        """Truncate ISO 8601 timestamp to seconds to absorb fractional-second format differences."""
        return ts[:19] if ts else ""

    to_process = []
    skipped    = 0
    for item in items:
        item_id     = item.get("id", "")
        sp_modified = item.get("last_modified_date", "")
        db_modified = db_map.get(item_id)

        if db_modified is not None and _norm_ts(db_modified) == _norm_ts(sp_modified) and sp_modified:
            _log.info(f"  [SKIP unchanged] {item.get('name', item_id)} — last_modified={_norm_ts(sp_modified)}")
            skipped += 1
        else:
            reason = "new" if db_modified is None else f"modified ({_norm_ts(db_modified)} → {_norm_ts(sp_modified)})"
            _log.info(f"  [QUEUE {reason}] {item.get('name', item_id)}")
            to_process.append(item)

    _log.info(f"Change filter: {len(to_process)} to process, {skipped} unchanged (skipped)")
    return to_process, skipped


def _download_and_vectorize(items: list[dict], logger: logging.Logger = None) -> tuple:
    _log  = logger or log
    ts    = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Track counts only — never accumulate full data in memory
    n_chunks  = 0
    n_files   = 0

    t_start        = time.time()
    skips          = []
    retries        = []
    ingest_results = []

    # Open output files once and write per-file in append/JSONL mode
    active_dir   = ts[:8]
    save_dir     = os.path.join(DOWNLOAD_PATH, active_dir)
    os.makedirs(save_dir, exist_ok=True)
    path_chunks  = os.path.join(save_dir, f"sp_chunks_{ts}.jsonl")
    path_raw     = os.path.join(save_dir, f"sp_files_chunk_{ts}.jsonl")
    path_vectors = os.path.join(save_dir, f"sp_vectors_{ts}.jsonl")

    _log.info(f"Starting vectorization of {len(items)} items")

    for idx, item in enumerate(items, start=1):
        if idx > 1:
            time.sleep(1.0)

        url      = item.get("BaseURL")
        filename = item.get("name", "")
        item_id  = item.get("id")

        print(f"[{idx}/{len(items)}] ── {filename} (item_id={item_id}, size={item.get('size','—')})")
        _log.info(f"[FILE {idx}/{len(items)}] START — filename={filename} item_id={item_id} size={item.get('size','—')}")

        if not url:
            print(f"  [STEP 1/6 DOWNLOAD] ERROR — No BaseURL, skipping")
            _log.error(f"  [STEP 1/6 DOWNLOAD] No BaseURL — skipping  item_id={item_id} filename={filename}")
            skips.append({"filename": filename, "size": item.get("size", "—"), "attempts": 0, "outcome": "no_url"})
            continue

        _attempts    = 0
        _outcome     = None
        MAX_ATTEMPTS = 3
        BACKOFF      = [5, 15, 30]
        content      = None

        for attempt in range(1, MAX_ATTEMPTS + 1):
            _attempts = attempt
            try:
                _log.info(f"  [STEP 1/6 DOWNLOAD] attempt {attempt}/{MAX_ATTEMPTS} — {url.split('?')[0]}")
                byte_chunks = []
                with requests.get(url, headers=_headers(), stream=True, timeout=(15, 300)) as r:
                    r.raise_for_status()
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            byte_chunks.append(chunk)
                content = b"".join(byte_chunks)
                dl_chunk_count = len(byte_chunks)
                del byte_chunks  # free HTTP chunk list — content holds the joined bytes
                print(f"  [STEP 1/6 DOWNLOAD] OK — {len(content) / 1024:.1f} KB ({dl_chunk_count} http chunks)")
                _log.info(f"  [STEP 1/6 DOWNLOAD] OK — {len(content) / 1024:.1f} KB ({dl_chunk_count} http chunks)")
                _outcome = "ok"
                break

            except requests.exceptions.ReadTimeout:
                print(f"  [STEP 1/6 DOWNLOAD] TIMEOUT attempt {attempt}/{MAX_ATTEMPTS}")
                _log.warning(f"  [STEP 1/6 DOWNLOAD] ReadTimeout — attempt {attempt}/{MAX_ATTEMPTS} filename={filename}")
                if attempt < MAX_ATTEMPTS:
                    time.sleep(BACKOFF[attempt - 1])
                else:
                    print(f"  [STEP 1/6 DOWNLOAD] TIMEOUT — all {MAX_ATTEMPTS} attempts exhausted, skipping")
                    _log.warning(f"  [STEP 1/6 DOWNLOAD] ReadTimeout — all {MAX_ATTEMPTS} attempts exhausted, skipping filename={filename}")
                    _outcome = "timeout_exhausted"

            except requests.exceptions.HTTPError as e:
                sc = e.response.status_code if e.response is not None else 0
                body = (e.response.text[:300] if e.response is not None else "(no response)")
                if sc == 429 and attempt < MAX_ATTEMPTS:
                    retry_after = int(e.response.headers.get("Retry-After", BACKOFF[attempt - 1]))
                    print(f"  [STEP 1/6 DOWNLOAD] 429 rate-limited — waiting {retry_after}s...")
                    _log.warning(f"  [STEP 1/6 DOWNLOAD] HTTP 429 Retry-After={retry_after}s filename={filename}")
                    time.sleep(retry_after)
                elif sc in (500, 502, 503, 504) and attempt < MAX_ATTEMPTS:
                    print(f"  [STEP 1/6 DOWNLOAD] HTTP {sc} — retrying in {BACKOFF[attempt - 1]}s...")
                    _log.warning(f"  [STEP 1/6 DOWNLOAD] HTTP {sc} attempt {attempt}/{MAX_ATTEMPTS} — retrying in {BACKOFF[attempt-1]}s  filename={filename}")
                    time.sleep(BACKOFF[attempt - 1])
                else:
                    print(f"  [STEP 1/6 DOWNLOAD] HTTP ERROR {sc} — {body}")
                    _log.error(f"  [STEP 1/6 DOWNLOAD] HTTPError {sc} — filename={filename}  body={body}", exc_info=True)
                    _outcome = f"http_{sc}"
                    break

            except Exception as e:
                print(f"  [STEP 1/6 DOWNLOAD] ERROR — {type(e).__name__}: {e}")
                _log.error(f"  [STEP 1/6 DOWNLOAD] Unexpected error — {type(e).__name__}: {e}  filename={filename}", exc_info=True)
                _outcome = "error"
                break

        if content is None:
            skips.append({"filename": filename, "size": item.get("size", "—"), "attempts": _attempts, "outcome": _outcome or "error"})
            continue

        if _attempts > 1:
            retries.append({"filename": filename, "size": item.get("size", "—"), "attempts": _attempts, "outcome": "recovered"})

        # Build raw binary chunks before freeing content
        _chunk_size = 8192
        _raw_chunk_dt = datetime.utcnow().isoformat()
        raw_chunk_records = []
        for _ci in range(0, len(content), _chunk_size):
            _raw = content[_ci:_ci + _chunk_size]
            raw_chunk_records.append({
                "item_id":        item_id,
                "filename":       filename,
                "chunk_datetime": _raw_chunk_dt,
                "chunk_size":     len(_raw),
                "chunk_id":       (_ci // _chunk_size) + 1,
                "chunk_data":     base64.b64encode(_raw).decode("utf-8"),
            })
        _log.info(f"  [STEP 2/6 EXTRACT] {len(raw_chunk_records)} raw binary chunks ready")

        _log.info(f"  [STEP 2/6 EXTRACT] Extracting text from {filename}")
        try:
            text = _extract_text(content, filename)
        except Exception as e:
            _log.error(f"  [STEP 2/6 EXTRACT] FAILED — {type(e).__name__}: {e}  filename={filename}", exc_info=True)
            skips.append({"filename": filename, "size": item.get("size", "—"), "attempts": _attempts, "outcome": "extract_error"})
            del content
            continue
        del content  # free download bytes immediately after extraction

        if not text.strip():
            print(f"  [STEP 2/6 EXTRACT] No text extracted — skipping")
            _log.warning(f"  [STEP 2/6 EXTRACT] Empty text — filename={filename}")
            skips.append({"filename": filename, "size": item.get("size", "—"), "attempts": _attempts, "outcome": "no_text"})
            continue
        _log.info(f"  [STEP 2/6 EXTRACT] OK — {len(text)} chars extracted")

        _log.info(f"  [STEP 3/6 CHUNK] Chunking text into 512-token segments")
        try:
            chunks = _semantic_chunks(text, max_tokens=512)
        except Exception as e:
            _log.error(f"  [STEP 3/6 CHUNK] FAILED — {type(e).__name__}: {e}  filename={filename}", exc_info=True)
            skips.append({"filename": filename, "size": item.get("size", "—"), "attempts": _attempts, "outcome": "chunk_error"})
            del text
            continue
        del text  # free extracted text after chunking
        print(f"  [STEP 3/6 CHUNK] {len(chunks)} semantic chunks")
        _log.info(f"  [STEP 3/6 CHUNK] OK — {len(chunks)} chunks")

        _log.info(f"  [STEP 4/6 EMBED] Encoding {len(chunks)} chunk(s) with all-MiniLM-L6-v2")
        file_chunks  = []
        file_vectors = []
        try:
            for i, chunk_text in enumerate(chunks, start=1):
                _log.debug(f"    Embedding chunk {i}/{len(chunks)}")
                embedding = _get_embed_model().encode(chunk_text).tolist()
                file_chunks.append({
                    "item_id":    item_id,
                    "filename":   filename,
                    "chunk_id":   i,
                    "chunk_data": chunk_text,
                })
                file_vectors.append({
                    "item_id":     item_id,
                    "filename":    filename,
                    "chunk_id":    i,
                    "vector_data": embedding,
                })
        except Exception as e:
            _log.error(f"  [STEP 4/6 EMBED] FAILED at chunk {i}/{len(chunks)} — {type(e).__name__}: {e}  filename={filename}", exc_info=True)
            skips.append({"filename": filename, "size": item.get("size", "—"), "attempts": _attempts, "outcome": "embed_error"})
            del chunks
            continue
        _log.info(f"  [STEP 4/6 EMBED] OK — {len(file_vectors)} vectors (384-dim)")

        file_raw = {
            "item_id":     item_id,
            "filename":    filename,
            "chunk_count": len(chunks),
        }

        _log.info(f"  [STEP 5/6 WRITE] Writing JSONL output files")
        try:
            _write_to_json(file_chunks,  path=path_chunks)
            _write_to_json(file_vectors, path=path_vectors)
            _write_to_json([file_raw],   path=path_raw)
            _log.info(f"  [STEP 5/6 WRITE] OK — chunks→{path_chunks}  vectors→{path_vectors}")
        except Exception as e:
            _log.error(f"  [STEP 5/6 WRITE] FAILED — {type(e).__name__}: {e}  filename={filename}", exc_info=True)

        _log.info(f"  [STEP 6/6 DB UPSERT] Ingesting into HANA — {len(file_chunks)} chunks, {len(file_vectors)} vectors")
        try:
            r1 = db.ingest(file_chunks, file_vectors, [item])
            _log.info(f"  [STEP 6/6 DB UPSERT] text+vectors OK — {r1}")
            try:
                db.ingest_raw_chunks(raw_chunk_records, source="sharepoint")
                _log.info(f"  [STEP 6/6 DB UPSERT] BRAIN_FILES_CHUNK OK — {len(raw_chunk_records)} raw chunks")
            except Exception as rc_err:
                _log.warning(f"  BRAIN_FILES_CHUNK upsert failed — {rc_err}")
            ingest_results.append({"filename": filename, "chunks": len(file_chunks), "status": "ok", **r1})
            try:
                db.ingest_file_registry([{
                    "item_id":          item_id,
                    "modified_datetime": item.get("last_modified_date", ""),
                    "filename":         filename,
                }], source="sharepoint")
            except Exception as reg_err:
                _log.warning(f"  FILE_REGISTRY upsert failed — {reg_err}")
        except Exception as db_err:
            ingest_results.append({"filename": filename, "chunks": len(file_chunks), "status": "error", "error": str(db_err)})
            _log.error(f"  [STEP 6/6 DB UPSERT] FAILED — {type(db_err).__name__}: {db_err}  filename={filename}", exc_info=True)

        n_chunks += len(file_chunks)
        n_files  += 1

        # Free per-file data — counts are tracked above
        del file_chunks, file_vectors, file_raw, chunks, raw_chunk_records

        print(f"  {n_chunks} total chunks so far")
        _log.info(f"  [FILE {idx}/{len(items)}] DONE — running total: {n_files} files, {n_chunks} chunks")

    print(f"\n[BATCH COMPLETE] {n_chunks} chunks, {n_files} files vectorized.")
    _log.info(f"Batch complete — {n_chunks} chunk records, {n_files} file records.")

    # path_chunks / path_raw / path_vectors were opened per-file above; save_dir already exists
    try:
        _db_stats = db.get_table_stats()
    except Exception as _e:
        _log.warning(f"Could not fetch DB table stats for analysis: {_e}")
        _db_stats = {}

    path_html = _generate_analysis_html(
        run_summary={
            "ts":               ts,
            "total_items":      len(items),
            "vectorized_count": n_files,
            "chunk_count":      n_chunks,
            "skips":            skips,
            "retries":          retries,
            "duration_s":       time.time() - t_start,
            "log_path":         _get_log_path(_log),
            "db_stats":         _db_stats,
            "file_sizes": {
                os.path.basename(path_chunks):  os.path.getsize(path_chunks)  if os.path.exists(path_chunks)  else 0,
                os.path.basename(path_raw):     os.path.getsize(path_raw)     if os.path.exists(path_raw)     else 0,
                os.path.basename(path_vectors): os.path.getsize(path_vectors) if os.path.exists(path_vectors) else 0,
            },
        },
        save_dir=save_dir,
    )
    path_log = _get_log_path(_log)

    # Flush log handler before uploading so the file is complete
    for _h in _log.handlers:
        _h.flush()
    s3_uploads = _upload_run_artifacts([path_html, path_log], logger=_log)
    if s3_uploads:
        _log.info(f"S3 artifacts uploaded: {list(s3_uploads.values())}")

    return n_chunks, n_files, path_chunks, path_raw, path_vectors, path_html, path_log, ingest_results


# ── S3 vectorization ──────────────────────────────────────────────────────────

def _fmt_bytes(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.2f} TB"


def _s3_obj_to_item(obj: dict) -> dict:
    """Convert an s3.list_objects() dict to the item shape expected by db.ingest()."""
    key      = obj["key"]
    filename = key.split("/")[-1] if "/" in key else key
    ext      = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    return {
        "id":                 obj.get("etag", key),
        "name":               filename,
        "type":               ext,
        "extension":          ext,
        "size":               _fmt_bytes(obj.get("size", 0)),
        "size_bytes":         obj.get("size", 0),
        "web_url":            "",
        "doc_path":           key,
        "last_modified_date": obj.get("last_modified", ""),
        "created_date":       "",
        "created_by":         "",
        "modified_by":        "",
        "shared":             "",
        "version_count":      0,
        "etag":               obj.get("etag", ""),
        "storage_class":      obj.get("storage_class", "STANDARD"),
    }


def _tmp_file_to_item(abs_path: str) -> dict:
    """Convert a local file path to the item shape expected by db.ingest()."""
    filename  = os.path.basename(abs_path)
    ext       = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    size      = os.path.getsize(abs_path) if os.path.exists(abs_path) else 0
    mtime     = os.path.getmtime(abs_path) if os.path.exists(abs_path) else 0
    mtime_iso = datetime.utcfromtimestamp(mtime).isoformat() + "Z" if mtime else ""
    item_id   = hashlib.md5(abs_path.encode()).hexdigest()
    return {
        "id":                 item_id,
        "name":               filename,
        "type":               ext,
        "extension":          ext,
        "size":               _fmt_bytes(size),
        "size_bytes":         size,
        "web_url":            "",
        "doc_path":           abs_path,
        "last_modified_date": mtime_iso,
        "created_date":       "",
        "created_by":         "",
        "modified_by":        "",
        "shared":             "",
        "version_count":      0,
        "etag":               item_id,
        "storage_class":      "LOCAL",
    }


def _download_and_vectorize_tmp(paths: list[str], logger: logging.Logger = None, skip_ingest: bool = False) -> tuple:
    """
    /tmp variant of _download_and_vectorize_s3.
    Reads each file from the local filesystem, then runs the same
    extract → chunk → embed → HANA upsert pipeline.

    Returns: (n_chunks, n_files, path_chunks, path_raw, path_vectors, path_root, path_html, path_log, ingest_results)
    """
    _log     = logger or log
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    n_chunks = 0
    n_files  = 0
    t_start  = time.time()
    skips    = []
    ingest_results = []

    active_dir   = ts[:8]
    save_dir     = os.path.join(DOWNLOAD_PATH, active_dir)
    os.makedirs(save_dir, exist_ok=True)
    path_chunks  = os.path.join(save_dir, f"tmp_chunks_{ts}.jsonl")
    path_raw     = os.path.join(save_dir, f"tmp_files_chunks_{ts}.json")
    path_vectors = os.path.join(save_dir, f"tmp_vectors_{ts}.jsonl")
    path_root    = os.path.join(save_dir, f"tmp_root_data_{ts}.json")

    root_data = []
    _log.info(f"TMP vectorization starting — {len(paths)} files")

    for idx, abs_path in enumerate(paths, start=1):
        filename = os.path.basename(abs_path)
        item     = _tmp_file_to_item(abs_path)
        item_id  = item["id"]

        print(f"[{idx}/{len(paths)}] ── {abs_path} ({item['size']})")
        _log.info(f"[TMP FILE {idx}/{len(paths)}] START — path={abs_path} size={item['size']}")

        # ── STEP 1: Read from local disk ──────────────────────────────────────
        try:
            with open(abs_path, "rb") as fh:
                content = fh.read()
            _log.info(f"  [STEP 1/5 READ] OK — {len(content) / 1024:.1f} KB")
        except Exception as e:
            print(f"  [STEP 1/5 READ] ERROR — {e}")
            _log.error(f"  [STEP 1/5 READ] FAILED — {type(e).__name__}: {e}  path={abs_path}", exc_info=True)
            skips.append({"filename": filename, "size": item["size"], "outcome": "read_error"})
            continue

        _chunk_size = 8192
        raw_chunk_records = []
        for _ci in range(0, len(content), _chunk_size):
            _raw_chunk = content[_ci:_ci + _chunk_size]
            raw_chunk_records.append({
                "item_id":        item_id,
                "filename":       filename,
                "chunk_datetime": datetime.now().isoformat(),
                "chunk_size":     len(_raw_chunk),
                "chunk_id":       (_ci // _chunk_size) + 1,
                "chunk_data":     base64.b64encode(_raw_chunk).decode("utf-8"),
            })
        _log.info(f"  [STEP 1/5 READ] {len(raw_chunk_records)} raw chunks base64-encoded")

        # ── STEP 2: Extract text ──────────────────────────────────────────────
        try:
            text = _extract_text(content, filename)
        except Exception as e:
            _log.error(f"  [STEP 2/5 EXTRACT] FAILED — {type(e).__name__}: {e}  path={abs_path}", exc_info=True)
            skips.append({"filename": filename, "size": item["size"], "outcome": "extract_error"})
            del content
            continue
        del content

        if not text.strip():
            _log.warning(f"  [STEP 2/5 EXTRACT] Empty text — path={abs_path}")
            skips.append({"filename": filename, "size": item["size"], "outcome": "no_text"})
            continue
        _log.info(f"  [STEP 2/5 EXTRACT] OK — {len(text)} chars")

        # ── STEP 3: Chunk ─────────────────────────────────────────────────────
        try:
            chunks = _semantic_chunks(text, max_tokens=512)
        except Exception as e:
            _log.error(f"  [STEP 3/5 CHUNK] FAILED — {type(e).__name__}: {e}  path={abs_path}", exc_info=True)
            skips.append({"filename": filename, "size": item["size"], "outcome": "chunk_error"})
            del text
            continue
        del text
        _log.info(f"  [STEP 3/5 CHUNK] {len(chunks)} chunks")

        # ── STEP 4: Embed ─────────────────────────────────────────────────────
        file_chunks  = []
        file_vectors = []
        try:
            for i, chunk_text in enumerate(chunks, start=1):
                embedding = _get_embed_model().encode(chunk_text).tolist()
                file_chunks.append({"item_id": item_id, "filename": filename, "chunk_id": i, "chunk_data": chunk_text})
                file_vectors.append({"item_id": item_id, "filename": filename, "chunk_id": i, "vector_data": embedding})
        except Exception as e:
            _log.error(f"  [STEP 4/5 EMBED] FAILED — {type(e).__name__}: {e}  path={abs_path}", exc_info=True)
            skips.append({"filename": filename, "size": item["size"], "outcome": "embed_error"})
            del chunks
            continue
        del chunks
        _log.info(f"  [STEP 4/5 EMBED] OK — {len(file_vectors)} vectors (384-dim)")

        # ── STEP 5: Write JSONL + upsert HANA ────────────────────────────────
        _write_to_json(file_chunks,       path=path_chunks)
        _write_to_json(file_vectors,      path=path_vectors)
        _write_to_json(raw_chunk_records, path=path_raw)

        if skip_ingest:
            ingest_results.append({"filename": filename, "path": abs_path, "chunks": len(file_chunks), "status": "skipped"})
            _log.info(f"  [STEP 5/5 DB UPSERT] SKIPPED (skip_ingest=True)")
        else:
            try:
                r1 = db.ingest(file_chunks, file_vectors, [item], source="tmp")
                _log.info(f"  [STEP 5/5 DB UPSERT] text+vectors OK — {r1}")
                try:
                    db.ingest_raw_chunks(raw_chunk_records, source="tmp")
                    _log.info(f"  [STEP 5/5 DB UPSERT] BRAIN_FILES_CHUNK OK — {len(raw_chunk_records)} raw chunks")
                except Exception as rc_err:
                    _log.warning(f"  BRAIN_FILES_CHUNK upsert failed — {rc_err}")
                ingest_results.append({"filename": filename, "path": abs_path, "chunks": len(file_chunks), "status": "ok", **r1})
                try:
                    db.ingest_file_registry([{
                        "item_id":           item_id,
                        "modified_datetime": item.get("last_modified_date", ""),
                        "filename":          filename,
                    }], source="tmp")
                except Exception as reg_err:
                    _log.warning(f"  FILE_REGISTRY upsert failed — {reg_err}")
            except Exception as db_err:
                ingest_results.append({"filename": filename, "path": abs_path, "chunks": len(file_chunks), "status": "error", "error": str(db_err)})
                _log.error(f"  [STEP 5/5 DB UPSERT] FAILED — {type(db_err).__name__}: {db_err}  path={abs_path}", exc_info=True)

        root_data.append({
            "item_id":       item_id,
            "filename":      filename,
            "chunk_count":   len(file_chunks),
            "path":          abs_path,
            "size":          item.get("size_bytes", 0),
            "last_modified": item.get("last_modified_date", ""),
        })

        n_chunks += len(file_chunks)
        n_files  += 1
        del file_chunks, file_vectors, raw_chunk_records
        _log.info(f"  [TMP FILE {idx}/{len(paths)}] DONE — running total: {n_files} files, {n_chunks} chunks")

    print(f"\n[TMP BATCH COMPLETE] {n_chunks} chunks, {n_files} files vectorized, {len(skips)} skipped.")
    _log.info(f"TMP batch complete — {n_chunks} chunks, {n_files} files, {len(skips)} skipped.")

    try:
        with open(path_root, "w", encoding="utf-8") as _rf:
            json.dump(root_data, _rf, ensure_ascii=False, indent=2)
        _log.info(f"Root data written — {len(root_data)} records → {path_root}")
    except Exception as _e:
        _log.error(f"Failed to write root data: {_e}", exc_info=True)

    try:
        _db_stats = db.get_table_stats()
    except Exception as _e:
        _db_stats = {}

    path_html = _generate_analysis_html(
        run_summary={
            "ts":               ts,
            "total_items":      len(paths),
            "vectorized_count": n_files,
            "chunk_count":      n_chunks,
            "skips":            skips,
            "retries":          [],
            "duration_s":       time.time() - t_start,
            "log_path":         _get_log_path(_log),
            "db_stats":         _db_stats,
            "file_sizes": {
                os.path.basename(path_chunks):  os.path.getsize(path_chunks)  if os.path.exists(path_chunks)  else 0,
                os.path.basename(path_raw):     os.path.getsize(path_raw)     if os.path.exists(path_raw)     else 0,
                os.path.basename(path_vectors): os.path.getsize(path_vectors) if os.path.exists(path_vectors) else 0,
                os.path.basename(path_root):    os.path.getsize(path_root)    if os.path.exists(path_root)    else 0,
            },
        },
        save_dir=save_dir,
    )
    path_log = _get_log_path(_log)
    for _h in _log.handlers:
        _h.flush()
    s3_uploads = _upload_run_artifacts([path_html, path_log], logger=_log)
    if s3_uploads:
        _log.info(f"S3 artifacts uploaded: {list(s3_uploads.values())}")

    return n_chunks, n_files, path_chunks, path_raw, path_vectors, path_root, path_html, path_log, ingest_results


def _download_and_vectorize_s3(objects: list[dict], logger: logging.Logger = None, skip_ingest: bool = False) -> tuple:
    """
    S3 variant of _download_and_vectorize.
    Downloads each object via s3.download_bytes(), then runs the same
    extract → chunk → embed → HANA upsert pipeline.

    Returns: (n_chunks, n_files, path_chunks, path_raw, path_vectors, path_html, path_log, ingest_results)
    """
    import s3 as s3_client

    _log     = logger or log
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    n_chunks = 0
    n_files  = 0
    t_start  = time.time()
    skips    = []
    ingest_results = []

    active_dir   = ts[:8]
    save_dir     = os.path.join(DOWNLOAD_PATH, active_dir)
    os.makedirs(save_dir, exist_ok=True)
    path_chunks  = os.path.join(save_dir, f"s3_chunks_{ts}.jsonl")
    path_raw     = os.path.join(save_dir, f"s3_files_chunks_{ts}.json")
    path_vectors = os.path.join(save_dir, f"s3_vectors_{ts}.jsonl")
    path_root    = os.path.join(save_dir, f"s3_root_data_{ts}.json")

    root_data = []
    _log.info(f"S3 vectorization starting — {len(objects)} objects")

    for idx, obj in enumerate(objects, start=1):
        key      = obj["key"]
        filename = key.split("/")[-1] if "/" in key else key
        item     = _s3_obj_to_item(obj)
        item_id  = item["id"]

        print(f"[{idx}/{len(objects)}] ── {key} ({item['size']})")
        _log.info(f"[S3 FILE {idx}/{len(objects)}] START — key={key} size={item['size']}")

        # ── STEP 1: Download from S3 ──────────────────────────────────────────
        try:
            content = s3_client.download_bytes(key)
            _log.info(f"  [STEP 1/5 DOWNLOAD] OK — {len(content) / 1024:.1f} KB")
        except Exception as e:
            print(f"  [STEP 1/5 DOWNLOAD] ERROR — {e}")
            _log.error(f"  [STEP 1/5 DOWNLOAD] FAILED — {type(e).__name__}: {e}  key={key}", exc_info=True)
            skips.append({"filename": key, "size": item["size"], "outcome": "download_error"})
            continue

        # Split raw bytes into 8192-byte chunks and base64-encode each one
        _chunk_size = 8192
        raw_chunk_records = []
        for _ci in range(0, len(content), _chunk_size):
            _raw_chunk = content[_ci:_ci + _chunk_size]
            raw_chunk_records.append({
                "item_id":        item_id,
                "filename":       filename,
                "chunk_datetime": datetime.now().isoformat(),
                "chunk_size":     len(_raw_chunk),
                "chunk_id":       (_ci // _chunk_size) + 1,
                "chunk_data":     base64.b64encode(_raw_chunk).decode("utf-8"),
            })
        _log.info(f"  [STEP 1/5 DOWNLOAD] {len(raw_chunk_records)} raw chunks base64-encoded")

        # ── STEP 2: Extract text ──────────────────────────────────────────────
        try:
            text = _extract_text(content, filename)
        except Exception as e:
            _log.error(f"  [STEP 2/5 EXTRACT] FAILED — {type(e).__name__}: {e}  key={key}", exc_info=True)
            skips.append({"filename": key, "size": item["size"], "outcome": "extract_error"})
            del content
            continue
        del content

        if not text.strip():
            _log.warning(f"  [STEP 2/5 EXTRACT] Empty text — key={key}")
            skips.append({"filename": key, "size": item["size"], "outcome": "no_text"})
            continue
        _log.info(f"  [STEP 2/5 EXTRACT] OK — {len(text)} chars")

        # ── STEP 3: Chunk ─────────────────────────────────────────────────────
        try:
            chunks = _semantic_chunks(text, max_tokens=512)
        except Exception as e:
            _log.error(f"  [STEP 3/5 CHUNK] FAILED — {type(e).__name__}: {e}  key={key}", exc_info=True)
            skips.append({"filename": key, "size": item["size"], "outcome": "chunk_error"})
            del text
            continue
        del text
        _log.info(f"  [STEP 3/5 CHUNK] {len(chunks)} chunks")

        # ── STEP 4: Embed ─────────────────────────────────────────────────────
        file_chunks  = []
        file_vectors = []
        try:
            for i, chunk_text in enumerate(chunks, start=1):
                embedding = _get_embed_model().encode(chunk_text).tolist()
                file_chunks.append({"item_id": item_id, "filename": filename, "chunk_id": i, "chunk_data": chunk_text})
                file_vectors.append({"item_id": item_id, "filename": filename, "chunk_id": i, "vector_data": embedding})
        except Exception as e:
            _log.error(f"  [STEP 4/5 EMBED] FAILED — {type(e).__name__}: {e}  key={key}", exc_info=True)
            skips.append({"filename": key, "size": item["size"], "outcome": "embed_error"})
            del chunks
            continue
        del chunks
        _log.info(f"  [STEP 4/5 EMBED] OK — {len(file_vectors)} vectors (384-dim)")

        # ── STEP 5: Write JSONL + upsert HANA ────────────────────────────────
        _write_to_json(file_chunks,        path=path_chunks)
        _write_to_json(file_vectors,       path=path_vectors)
        _write_to_json(raw_chunk_records,  path=path_raw)

        if skip_ingest:
            ingest_results.append({"filename": filename, "key": key, "chunks": len(file_chunks), "status": "skipped"})
            _log.info(f"  [STEP 5/5 DB UPSERT] SKIPPED (skip_ingest=True)")
        else:
            try:
                r1 = db.ingest(file_chunks, file_vectors, [item], source="s3")
                _log.info(f"  [STEP 5/5 DB UPSERT] text+vectors OK — {r1}")
                try:
                    db.ingest_raw_chunks(raw_chunk_records, source="s3")
                    _log.info(f"  [STEP 5/5 DB UPSERT] BRAIN_FILES_CHUNK OK — {len(raw_chunk_records)} raw chunks")
                except Exception as rc_err:
                    _log.warning(f"  BRAIN_FILES_CHUNK upsert failed — {rc_err}")
                ingest_results.append({"filename": filename, "key": key, "chunks": len(file_chunks), "status": "ok", **r1})
                try:
                    db.ingest_file_registry([{
                        "item_id":          item_id,
                        "modified_datetime": item.get("last_modified_date", ""),
                        "filename":         filename,
                    }], source="s3")
                except Exception as reg_err:
                    _log.warning(f"  FILE_REGISTRY upsert failed — {reg_err}")
            except Exception as db_err:
                ingest_results.append({"filename": filename, "key": key, "chunks": len(file_chunks), "status": "error", "error": str(db_err)})
                _log.error(f"  [STEP 5/5 DB UPSERT] FAILED — {type(db_err).__name__}: {db_err}  key={key}", exc_info=True)

        root_data.append({
            "item_id":       item_id,
            "filename":      filename,
            "chunk_count":   len(file_chunks),
            "key":           key,
            "size":          obj.get("size", 0),
            "last_modified": obj.get("last_modified", ""),
            "etag":          obj.get("etag", ""),
            "storage_class": obj.get("storage_class", "STANDARD"),
        })

        n_chunks += len(file_chunks)
        n_files  += 1
        del file_chunks, file_vectors, raw_chunk_records
        _log.info(f"  [S3 FILE {idx}/{len(objects)}] DONE — running total: {n_files} files, {n_chunks} chunks")

    print(f"\n[S3 BATCH COMPLETE] {n_chunks} chunks, {n_files} files vectorized, {len(skips)} skipped.")
    _log.info(f"S3 batch complete — {n_chunks} chunks, {n_files} files, {len(skips)} skipped.")

    try:
        with open(path_root, "w", encoding="utf-8") as _rf:
            json.dump(root_data, _rf, ensure_ascii=False, indent=2)
        _log.info(f"Root data written — {len(root_data)} records → {path_root}")
    except Exception as _e:
        _log.error(f"Failed to write root data: {_e}", exc_info=True)

    try:
        _db_stats = db.get_table_stats()
    except Exception as _e:
        _db_stats = {}

    path_html = _generate_analysis_html(
        run_summary={
            "ts":               ts,
            "total_items":      len(objects),
            "vectorized_count": n_files,
            "chunk_count":      n_chunks,
            "skips":            skips,
            "retries":          [],
            "duration_s":       time.time() - t_start,
            "log_path":         _get_log_path(_log),
            "db_stats":         _db_stats,
            "file_sizes": {
                os.path.basename(path_chunks):  os.path.getsize(path_chunks)  if os.path.exists(path_chunks)  else 0,
                os.path.basename(path_raw):     os.path.getsize(path_raw)     if os.path.exists(path_raw)     else 0,
                os.path.basename(path_vectors): os.path.getsize(path_vectors) if os.path.exists(path_vectors) else 0,
                os.path.basename(path_root):    os.path.getsize(path_root)    if os.path.exists(path_root)    else 0,
            },
        },
        save_dir=save_dir,
    )
    path_log = _get_log_path(_log)
    for _h in _log.handlers:
        _h.flush()
    s3_uploads = _upload_run_artifacts([path_html, path_log], logger=_log)
    if s3_uploads:
        _log.info(f"S3 artifacts uploaded: {list(s3_uploads.values())}")

    return n_chunks, n_files, path_chunks, path_raw, path_vectors, path_root, path_html, path_log, ingest_results


# ── Lifespan ──────────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    global site_id, drive_id

    # Validate credentials source before attempting any network call
    if _ON_CF:
        try:
            _get_dest_svc_creds()
            log.info(f"CF mode — MS Graph credentials via BTP Destination '{DESTINATION_NAME}'")
        except RuntimeError as e:
            log.error(str(e))
            raise
    else:
        missing = [v for v in ("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET") if not os.getenv(v)]
        if missing:
            log.error(f"Missing required env vars: {', '.join(missing)}")
            raise RuntimeError(f"Missing required env vars: {', '.join(missing)}")

    log.info("Connecting to SharePoint...")
    try:
        site_id, drive_id = await asyncio.to_thread(_get_SiteDrive_ID)
        log.info(f"Connected — site_id={site_id}  drive_id={drive_id}")
    except Exception as e:
        log.warning(f"SharePoint init failed (non-fatal): {e} — endpoints will return 503 until connection is available")

    # Tables are managed by the HDI module (db/src/*.hdbtable) — not created here.
    # try:
    #     table_status = await asyncio.to_thread(db.create_tables)
    #     for table, result in table_status.items():
    #         log.info(f"  DB table {table}: {result}")
    # except Exception as e:
    #     log.warning(f"HANA table setup failed (non-fatal): {e}")

    # S3 / Object Store — non-fatal; warn and continue if not configured
    try:
        s3_info = await asyncio.to_thread(s3.ping)
        log.info(f"S3 connected — bucket={s3_info['bucket']} region={s3_info['region']}")
    except Exception as e:
        log.warning(f"S3 not available (non-fatal): {e} — /s3/* endpoints will return 503")

    yield
    log.info("Shutdown complete.")


# ── App ───────────────────────────────────────────────────────────────────────

from zoneinfo import ZoneInfo as _ZoneInfo
_now_utc    = datetime.utcnow()
_now_manila = datetime.now(tz=_ZoneInfo("Asia/Manila"))
_SERVER_START = (
    f"{_now_utc.strftime('%Y-%m-%d %H:%M:%S')} UTC"
    f"  |  {_now_manila.strftime('%Y-%m-%d %H:%M:%S')} PHT (Asia/Manila)"
)

_DESCRIPTION = f"""
## SharePoint RAG API

> 🕐 **Server started:** `{_SERVER_START}`

A FastAPI service deployed on **SAP BTP Cloud Foundry** that connects to a Microsoft SharePoint
document library, extracts text from files, chunks and embeds the content into semantic vectors,
and stores everything in **SAP HANA Cloud** for retrieval-augmented generation (RAG) queries.

---

### Architecture

```
SharePoint (MS Graph API)
        │
        ▼
   api.py  ──► text extraction (docx, xlsx, pdf, pptx, csv, …)
        │      semantic chunking (512-token windows)
        │      embedding (all-MiniLM-L6-v2, 384-dim)
        ▼
   db.py   ──► SAP HANA Cloud  (BRAIN_SP_FILES · BRAIN_TEXT_CHUNKS · BRAIN_VECTORS · BRAIN_SP_FILES_VERSIONS · BRAIN_SP_FILES_CHUNK)
        │
        ▼
   s3.py   ──► SAP BTP Object Store / S3  (run artifacts: HTML reports, logs)
```

---

### Authentication

| Environment | Token source |
|---|---|
| **SAP BTP (CF)** | BTP Destination Service → destination `BTP-CP_MS_Graph` → Azure AD OAuth2 client credentials |
| **Local dev** | `TENANT_ID` / `CLIENT_ID` / `CLIENT_SECRET` env vars → Azure AD directly |

Tokens are cached and auto-refreshed 60 s before expiry.

---

### Supported file types

`docx` · `dotx` · `doc` · `xlsx` · `xlsm` · `xls` · `pdf` · `pptx` · `pptm` · `ppt` ·
`rtf` · `odt` · `eml` · `msg` · `csv` · `json` · `xml` · `html` · `htm` · `zip` (recursive) ·
`txt` · `md` · `py` · `js` · `ts` · `sql` · `yaml` · `yml` · `ini` · `toml` · `log` ·
`cfg` · `java` · `cs` · `abap`

---

### Change detection

Every vectorize call compares the SharePoint `lastModifiedDateTime` against the value stored in
`BRAIN_SP_FILES`. Files whose timestamp has not changed are skipped — only new or modified files are
downloaded and re-vectorized. The `skipped_unchanged` field in every response shows the count.

---

### Output files (written to `DOWNLOAD_PATH/<YYYYMMDD>/`)

| File | Contents |
|---|---|
| `sp_chunks_<ts>.jsonl` | Text chunks (item_id, chunk_id, chunk_data, filename) |
| `sp_vectors_<ts>.jsonl` | Embedding vectors (item_id, chunk_id, vector_data) |
| `sp_files_chunk_<ts>.jsonl` | Raw download chunks (base64, for rebuild) |
| `sp_root_data_<ts>.json` | Flat file metadata array |
| `sp_files_versions_<ts>.json` | Version history array |
| `sp_analysis_<ts>.html` | Run summary report (auto-uploaded to S3) |
| `<endpoint>_<ts>.log` | Per-call structured log (auto-uploaded to S3) |

---

### Tags

| Tag | Description |
|---|---|
| **system** | Health checks, connectivity diagnostics, /tmp browser |
| **files** | SharePoint file listing, metadata sync to HANA, download, rebuild |
| **rag** | Full vectorize pipeline — download → extract → chunk → embed → upsert |
| **analysis** | View HTML run reports and logs generated by vectorize runs |
| **storage** | Browse and manage local `/tmp` output files |
| **s3** | SAP BTP Object Store / S3 — list, upload, download, delete, presigned URLs |
| **database** | HANA vector search, CRUD, table management (`/db/*` routes) |
"""

app = FastAPI(
    title       = "SharePoint RAG API",
    version     = "1.0.0",
    description = _DESCRIPTION,
    lifespan    = lifespan,
    openapi_tags = [
        {"name": "system",   "description": "Health checks and connectivity diagnostics for SharePoint, HANA, S3, and BTP Destination Service."},
        {"name": "files",    "description": "List SharePoint files, sync metadata to HANA, download file bytes, and rebuild files from stored chunks."},
        {"name": "rag",      "description": "End-to-end vectorization pipeline: fetch → extract text → semantic chunk → embed → upsert to HANA. Supports root, folder, single-item, and POST-body modes."},
        {"name": "analysis", "description": "View auto-generated HTML run reports and per-call log files produced by vectorize runs."},
        {"name": "storage",  "description": "Browse, download, and delete files written to the local DOWNLOAD_PATH (/tmp on CF)."},
        {"name": "s3",       "description": "SAP BTP Object Store / S3 operations: list objects, upload, download, delete, presigned URLs."},
        {"name": "database", "description": "SAP HANA CRUD and vector similarity search across BRAIN_* tables."},
    ],
)

app.add_middleware(
    CORSMiddleware,
    allow_origins  = ["*"],
    allow_methods  = ["*"],
    allow_headers  = ["*"],
)

app.include_router(db.router)


# ── Dependency ────────────────────────────────────────────────────────────────

def require_connection():
    if drive_id is None:
        raise HTTPException(status_code=503, detail="SharePoint connection not ready")


def require_destructive():
    if os.getenv("ENABLE_DESTRUCTIVE", "false").lower() != "true":
        raise HTTPException(
            status_code=403,
            detail="Destructive actions are disabled on this instance. "
                   "Set ENABLE_DESTRUCTIVE=true and restart the app to enable."
        )


def require_debug():
    if os.getenv("ENABLE_DEBUG", "false").lower() != "true":
        raise HTTPException(
            status_code=403,
            detail="Debug endpoints are disabled on this instance. "
                   "Set ENABLE_DEBUG=true and restart the app to enable."
        )


# ── Models ────────────────────────────────────────────────────────────────────

class FilesResponse(BaseModel):
    count:     int        = Field(..., description="Total number of files returned from SharePoint.")
    files:     list[dict] = Field(..., description="Flattened file metadata array. Each item contains: id, name, type, extension, size, size_bytes, last_modified_date, created_date, created_by, modified_by, web_url, doc_path, child_count, versionURL, BaseURL, shared.")
    db_ingest: dict | None = Field(None, description="HANA upsert result — files/chunks/vectors counts, or an error message.")

class VectorizeRequest(BaseModel):
    items:    list[dict] | None = Field(None, description="Explicit list of SharePoint item dicts to vectorize (as returned by /files or /files/all). Omit to auto-fetch all root files with version history via the delta API.")
    save_dir: str        | None = Field(None, description="Absolute path for output JSONL/JSON files. Defaults to DOWNLOAD_PATH/<YYYYMMDD>/.")

class S3VectorizeRequest(BaseModel):
    prefix:     str       = Field(...,  description="S3 folder prefix to scan, e.g. 'BTP-Catalyst/' or 'knowledge-base/docs/'")
    force:      bool      = Field(False, description="When true, re-vectorize all files even if unchanged in HANA")
    extensions: list[str] = Field(
        default_factory=list,
        description="Whitelist of file extensions to process (lowercase, no dot). Empty list = all supported types.",
    )

class S3VectorizeItemRequest(BaseModel):
    key:   str  = Field(...,  description="Full S3 object key, e.g. 'BTP-Catalyst/docs/report.pdf'")
    force: bool = Field(False, description="When true, re-vectorize even if the file is already up-to-date in HANA")

class S3VectorizeItemResponse(BaseModel):
    key:               str        = Field(..., description="The S3 object key that was processed.")
    filename:          str        = Field(..., description="Filename portion of the key.")
    count_chunks:      int        = Field(..., description="Number of semantic chunks embedded and upserted. 0 if skipped.")
    skipped_unchanged: int        = Field(0,   description="1 if skipped because already up-to-date in HANA, 0 if processed.")
    saved_to:          list[str]  = Field(..., description="Local paths written: chunks, vectors, raw-chunk JSONL, HTML report, call log.")
    db_ingest:         dict | None = Field(None, description="HANA upsert result for this file.")

class VectorizeResponse(BaseModel):
    count_chunks:      int       = Field(..., description="Total semantic chunk records embedded and upserted to HANA in this run.")
    count_files:       int       = Field(..., description="Number of files successfully downloaded, extracted, and vectorized.")
    skipped_unchanged: int       = Field(0,   description="Files skipped because their SharePoint lastModifiedDateTime already matched the value stored in HANA.")
    saved_to:          list[str] = Field(..., description="Absolute local paths of output files written: sp_chunks, sp_vectors, sp_files_chunk JSONL, HTML report, call log.")
    db_ingest:         dict | None = Field(None, description="Per-file HANA upsert results (chunks, vectors, files counts) and optional version ingest counts.")

class PipelineRequest(BaseModel):
    save_dir: str | None = Field(None, description="Absolute path for output files. Defaults to DOWNLOAD_PATH/<YYYYMMDD>/.")
    delay:    float      = Field(1.0,  ge=0.0, le=10.0, description="Seconds to wait between SharePoint version-fetch requests to avoid Graph API throttling. Range: 0–10 s.")

class PipelineResponse(BaseModel):
    count_root_files:  int       = Field(..., description="Total files discovered across the entire SharePoint root drive delta.")
    count_chunks:      int       = Field(..., description="Total semantic chunk records embedded and upserted to HANA.")
    count_files:       int       = Field(..., description="Number of files that were downloaded and vectorized (unchanged files are excluded).")
    skipped_unchanged: int       = Field(0,   description="Files whose lastModifiedDateTime matched the value in HANA — no re-processing needed.")
    saved_to:          list[str] = Field(..., description="Absolute local paths of all output files: chunks, vectors, raw-chunks, root data JSON, versions JSON, HTML report, call log.")
    db_ingest:         dict | None = Field(None, description="Per-file HANA upsert results plus version ingest counts.")

class DownloadRequest(BaseModel):
    item:     dict       = Field(..., description="A SharePoint file metadata dict (from /files or /files/all). Must include 'BaseURL' (Graph content URL) and 'name' (filename).")
    save_dir: str | None = Field(None, description="Directory to save the downloaded file and its chunk JSONL. Defaults to DOWNLOAD_PATH/<today>/.")

class RebuildRequest(BaseModel):
    chunks:   list[dict] = Field(..., description="List of chunk dicts from files_list_chunk.json. Each must have: chunk_id (int), chunk_data (base64 string), filename (str).")
    save_dir: str | None  = Field(None, description="Directory to write the rebuilt file to disk. If omitted the reconstructed bytes are only streamed back — not saved.")

class VectorizeItemRequest(BaseModel):
    item_id: str = Field(..., description="SharePoint Drive item ID (opaque string from the Graph API, e.g. '01ABCDEF...'). Obtain from GET /files, /files/all, or /files/search.")

class VectorizeItemResponse(BaseModel):
    item_id:           str       = Field(..., description="The SharePoint item ID that was processed.")
    filename:          str       = Field(..., description="Display name of the file (e.g. 'Report Q1 2026.pdf').")
    count_chunks:      int       = Field(..., description="Number of semantic chunks embedded and upserted for this file. 0 if skipped.")
    skipped_unchanged: int       = Field(0,   description="1 if the file was skipped because it is already up-to-date in HANA, 0 if it was processed.")
    saved_to:          list[str] = Field(..., description="Local paths written: chunks JSONL, vectors JSONL, raw-chunk JSONL, HTML report, call log.")
    db_ingest:         dict | None = Field(None, description="HANA upsert result for this file's chunks, vectors, and version records.")

class S3UploadStorageRequest(BaseModel):
    day:          str      = Field(..., description="Date folder name under DOWNLOAD_PATH (e.g. '20260829'). Matches the 'date' field from GET /storage.")
    filename:     str      = Field(..., description="Filename within that date folder (e.g. 'sp_chunks_20260829_143000.jsonl').")
    s3_key:       str | None = Field(None, description="Destination S3 object key. Defaults to '<day>/<filename>' if not supplied.")
    content_type: str        = Field("application/octet-stream", description="MIME content type for the S3 object (e.g. 'application/json', 'text/html').")

class S3BulkDeleteRequest(BaseModel):
    keys: list[str] = Field(..., description="List of S3 object keys to delete. Up to 1 000 keys per call (S3 API limit).")

class TmpVectorizeRequest(BaseModel):
    folder:     str       = Field("/tmp", description="Absolute path of the local folder to scan, e.g. '/tmp/uploads' or '/tmp/docs'.")
    force:      bool      = Field(False,  description="When true, re-vectorize all files even if unchanged in HANA.")
    extensions: list[str] = Field(
        default_factory=list,
        description="Whitelist of file extensions to process (lowercase, no dot). Empty list = all supported types.",
    )

class TmpVectorizeItemRequest(BaseModel):
    path:  str  = Field(...,  description="Absolute path of the local file to vectorize, e.g. '/tmp/uploads/report.pdf'.")
    force: bool = Field(False, description="When true, re-vectorize even if the file is already up-to-date in HANA.")

class S3PresignedUploadRequest(BaseModel):
    s3_key:       str = Field(..., description="Destination S3 object key for the upload.")
    content_type: str = Field("application/octet-stream", description="MIME type the uploading client must declare in their Content-Type header.")
    expiry:       int = Field(3600, ge=60, le=86400, description="Presigned URL lifetime in seconds. Min 60 s, max 86 400 s (24 h).")
    max_size_mb:  int = Field(100,  ge=1,  le=5120,  description="Maximum allowed upload size in MB, enforced via an S3 content-length-range policy condition.")


# ── Routes ────────────────────────────────────────────────────────────────────

@app.get("/health", tags=["system"], summary="Application health check")
def health():
    """
    Returns the current health status of the API and confirms that the SharePoint
    connection was established at startup.

    **Response fields**
    - `status` — always `"ok"` if the app is running
    - `site_id` — SharePoint site ID resolved at startup (null if connection failed)
    - `drive_id` — SharePoint document library drive ID (null if connection failed)

    Used by the CF health-check probe (`health-check-http-endpoint: /health`).
    If `drive_id` is null the app is running but all SharePoint endpoints will return 503.
    """
    clog = _call_logger("health")
    clog.info("GET /health called")
    result = {"status": "ok", "site_id": site_id, "drive_id": drive_id}
    clog.info(f"Response: {result}")
    return result


@app.get("/debug/destination", tags=["system"], summary="Diagnose BTP Destination Service connectivity", dependencies=[Depends(require_debug)])
def debug_destination():
    """
    Step-by-step diagnostic for the BTP Destination Service → MS Graph token chain.
    Only meaningful when running on Cloud Foundry (`VCAP_SERVICES` is present).

    **Steps checked**
    1. `step_binding` — reads the Destination Service VCAP_SERVICES binding
    2. `step_xsuaa_token` — obtains an XSUAA OAuth2 token for the Destination Service API
    3. `step_dest_api` — calls the Destination Service API to fetch destination `BTP-CP_MS_Graph`; returns the config (clientSecret stripped)
    4. `step_ms_graph_token` — fetches an MS Graph Bearer token end-to-end and reports its expiry

    **Local dev** — returns `{"on_cf": false}` immediately (no VCAP_SERVICES present).
    """
    if not _ON_CF:
        return {"on_cf": False, "message": "Not running on CF — destination service not applicable"}

    result: dict = {"on_cf": True, "destination_name": DESTINATION_NAME}

    # Step 1: get Destination Service binding
    try:
        creds = _get_dest_svc_creds()
        result["dest_svc_uri"]      = creds.get("uri")
        result["dest_svc_url"]      = creds.get("url")
        result["dest_svc_clientid"] = creds.get("clientid")
        result["step_binding"]      = "ok"
    except Exception as e:
        result["step_binding"] = f"FAILED: {e}"
        return result

    # Step 2: get XSUAA token for Destination Service
    try:
        dest_token = _get_dest_svc_token()
        result["step_xsuaa_token"] = "ok"
    except Exception as e:
        result["step_xsuaa_token"] = f"FAILED: {e}"
        return result

    # Step 3: call Destination Service API for the named destination
    try:
        resp = requests.get(
            f"{creds['uri']}/destination-configuration/v1/destinations/{DESTINATION_NAME}",
            headers={"Authorization": f"Bearer {dest_token}"},
            timeout=15,
        )
        result["step_dest_api_status"] = resp.status_code
        if resp.status_code == 200:
            payload = resp.json()
            config  = payload.get("destinationConfiguration", {})
            # Strip secret before returning
            config_safe = {k: v for k, v in config.items() if "secret" not in k.lower() and "Secret" not in k}
            result["step_dest_api"]       = "ok"
            result["destination_config"]  = config_safe
            result["auth_tokens_present"] = [
                {"type": t.get("type"), "has_value": bool(t.get("value"))}
                for t in payload.get("authTokens", [])
            ]
        else:
            result["step_dest_api"] = f"FAILED: HTTP {resp.status_code}"
            result["response_body"] = resp.text[:500]
    except Exception as e:
        result["step_dest_api"] = f"FAILED: {e}"

    # Step 4: try getting an MS Graph token end-to-end
    try:
        token, expires_in = _get_ms_graph_token_from_destination()
        result["step_ms_graph_token"] = f"ok — expires_in={expires_in}s, token_prefix={token[:20]}..."
    except Exception as e:
        result["step_ms_graph_token"] = f"FAILED: {e}"

    return result


@app.get("/debug/db", tags=["system"], summary="Diagnose SAP HANA connectivity", dependencies=[Depends(require_debug)])
def debug_db():
    """
    Tests the HANA Cloud connection and returns basic connection info.

    **Response fields**
    - `host` / `port` / `user` / `schema` — resolved from VCAP_SERVICES, hana_creds.json, or env vars
    - `connected` — true if the TCP handshake and login succeeded
    - `db_version` — HANA database version string from `SYS.M_DATABASE`
    - `step_connect` — `"ok"` or `"FAILED: <error>"` for quick triage

    Credential resolution order: `VCAP_SERVICES` → `hana_creds.json` → env vars (`HANA_HOST` / `HANA_PORT` / `HANA_USER` / `HANA_PASSWORD`).
    """
    return db.connection_info()


@app.get("/analysis", tags=["analysis"], summary="List all run reports and logs")
def list_analysis():
    """
    Lists every vectorize analysis HTML report and per-call log file found under `DOWNLOAD_PATH`,
    grouped by date folder, newest first.

    **Response fields**
    - `count` — total number of report/log files found
    - `reports` — array of `{date, type, filename, path, url}` entries where `type` is `"analysis"` (HTML) or `"log"`
    - `url` — relative URL to view the file via `GET /analysis/{day}/{filename}`

    Reports are generated automatically at the end of every `/vectorize`, `/vectorize/folder`,
    `/vectorize/item`, and `/pipeline` run. They are also uploaded to S3 under `runs/<YYYYMMDD>/`.
    """
    results = []
    try:
        day_dirs = sorted(os.listdir(DOWNLOAD_PATH), reverse=True)
    except FileNotFoundError:
        return {"count": 0, "reports": []}
    for day_dir in day_dirs:
        day_path = os.path.join(DOWNLOAD_PATH, day_dir)
        if not os.path.isdir(day_path):
            continue
        for fname in sorted(os.listdir(day_path), reverse=True):
            if fname.startswith("sp_analysis_") and fname.endswith(".html"):
                results.append({"date": day_dir, "type": "analysis", "filename": fname,
                                 "path": os.path.join(day_path, fname), "url": f"/analysis/{day_dir}/{fname}"})
            elif fname.endswith(".log"):
                results.append({"date": day_dir, "type": "log", "filename": fname,
                                 "path": os.path.join(day_path, fname), "url": f"/analysis/{day_dir}/{fname}"})
    return {"count": len(results), "reports": results}


@app.get("/analysis/latest", tags=["analysis"], response_class=HTMLResponse, summary="View the most recent run report")
def analysis_latest():
    """
    Renders the most recent vectorize analysis HTML report directly in the browser.

    Scans `DOWNLOAD_PATH` date folders in reverse chronological order and returns the first
    `sp_analysis_*.html` file found. Raises **404** if no runs have been completed yet.

    Open this URL in a browser after any vectorize run to review:
    - files vectorized vs skipped
    - chunk counts and coverage
    - retry and timeout details
    - HANA table row counts
    - output file sizes
    """
    try:
        day_dirs = sorted(os.listdir(DOWNLOAD_PATH), reverse=True)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="No analysis reports found. Run /vectorize first.")
    for day_dir in day_dirs:
        day_path = os.path.join(DOWNLOAD_PATH, day_dir)
        if not os.path.isdir(day_path):
            continue
        for fname in sorted(os.listdir(day_path), reverse=True):
            if fname.startswith("sp_analysis_") and fname.endswith(".html"):
                with open(os.path.join(day_path, fname), encoding="utf-8") as fh:
                    return HTMLResponse(content=fh.read())
    raise HTTPException(status_code=404, detail="No analysis reports found. Run /vectorize first.")


@app.get("/analysis/db", tags=["analysis"], response_class=HTMLResponse, summary="Live DB metrics dashboard — all HANA tables")
def analysis_db():
    """
    Renders a live HTML dashboard with metrics collected from all five BRAIN tables.

    **Sections**
    - **Overview** — total files, chunks, vectors, vector coverage %, total raw size, avg chunk length
    - **Health Alerts** — files missing a chunk summary; chunks without a matching vector
    - **File Types** — count and total size per file extension, with distribution bar
    - **Shared vs Private** — breakdown of the `shared` column
    - **Indexing Timeline** — files indexed per day (bar chart)
    - **Top 10 by Chunk Count** — heaviest content files
    - **Recently Indexed** — last 10 files added
    - **Version History** — total records, avg versions/file, top 10 most-versioned files
    - **Contributors** — top 5 modifiers and creators from SharePoint metadata

    Returns **503** if HANA is unreachable.
    """
    try:
        from db import get_db_analysis
        data = get_db_analysis()
    except Exception as e:
        clog.error(f"DB analysis failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=503, detail=f"HANA unavailable: {e}")
    return HTMLResponse(content=_generate_db_analysis_html(data))


@app.get("/analysis/{day}/{filename}", tags=["analysis"], response_class=HTMLResponse, summary="View a specific report or log by date and filename")
def analysis_by_name(day: str, filename: str):
    """
    Returns a specific analysis HTML report or log file from a dated output folder.

    - **HTML files** (`.html`) — rendered as-is in the browser
    - **Log files** (`.log`) — rendered as a styled monospace page in the browser

    Path is validated against `DOWNLOAD_PATH` to prevent directory traversal.
    Use the `url` field from `GET /analysis` to build the correct path.
    """
    safe_root = os.path.realpath(DOWNLOAD_PATH)
    resolved  = os.path.realpath(os.path.join(DOWNLOAD_PATH, day, filename))
    if not resolved.startswith(safe_root + os.sep):
        raise HTTPException(status_code=400, detail="Invalid path")
    if not os.path.isfile(resolved):
        raise HTTPException(status_code=404, detail="File not found")
    if filename.endswith(".html"):
        with open(resolved, encoding="utf-8") as fh:
            return HTMLResponse(content=fh.read())
    if filename.endswith(".log"):
        with open(resolved, encoding="utf-8") as fh:
            raw = fh.read()
        safe_name = html.escape(filename)
        safe_body = html.escape(raw)
        return HTMLResponse(content=(
            f"<!DOCTYPE html><html><head><meta charset='UTF-8'/>"
            f"<title>{safe_name}</title></head>"
            f"<body style='background:#0f1117;color:#e2e8f0;font-family:monospace;padding:2rem'>"
            f"<h2 style='color:#94a3b8;margin-bottom:1rem'>{safe_name}</h2>"
            f"<pre style='font-size:.82rem;line-height:1.6;white-space:pre-wrap'>{safe_body}</pre>"
            f"</body></html>"
        ))
    raise HTTPException(status_code=400, detail="Only .html and .log files are supported")


# ── Storage — file manager for DOWNLOAD_PATH (/tmp) ──────────────────────────

def _safe_resolve(day: str, filename: str) -> str:
    """Resolve and validate path stays within DOWNLOAD_PATH. Raises 400 on traversal."""
    safe_root = os.path.realpath(DOWNLOAD_PATH)
    resolved  = os.path.realpath(os.path.join(DOWNLOAD_PATH, day, filename))
    if not resolved.startswith(safe_root + os.sep):
        raise HTTPException(status_code=400, detail="Invalid path")
    return resolved


@app.get("/storage", tags=["storage"], summary="List all local output files by date folder")
def storage_list():
    """
    Lists every file written to `DOWNLOAD_PATH` (typically `/tmp` on CF), grouped by date folder,
    newest folder first.

    **Response fields**
    - `count` — total file count across all folders
    - `total_size` / `total_size_bytes` — combined size of all files
    - `folders` — array of `{date, file_count, files[]}` where each file has `filename`, `size`, `modified_at`, `download_url`, `delete_url`

    **Note:** On SAP BTP CF, `/tmp` is ephemeral and wiped on dyno restart. Use `POST /s3/upload-storage-folder/{day}` to persist important files to S3 before that happens.
    """
    result = []
    try:
        day_dirs = sorted(os.listdir(DOWNLOAD_PATH), reverse=True)
    except FileNotFoundError:
        return {"count": 0, "total_size_bytes": 0, "folders": []}

    total_bytes = 0
    for day_dir in day_dirs:
        day_path = os.path.join(DOWNLOAD_PATH, day_dir)
        if not os.path.isdir(day_path):
            continue
        files = []
        for fname in sorted(os.listdir(day_path)):
            fpath = os.path.join(day_path, fname)
            if not os.path.isfile(fpath):
                continue
            size = os.path.getsize(fpath)
            total_bytes += size
            files.append({
                "filename":    fname,
                "size_bytes":  size,
                "size":        _fmt_size(size),
                "modified_at": datetime.fromtimestamp(os.path.getmtime(fpath)).isoformat(),
                "download_url": f"/storage/{day_dir}/{fname}",
                "delete_url":   f"/storage/{day_dir}/{fname}",
            })
        if files:
            result.append({"date": day_dir, "file_count": len(files), "files": files})

    return {
        "count":            sum(f["file_count"] for f in result),
        "total_size_bytes": total_bytes,
        "total_size":       _fmt_size(total_bytes),
        "folders":          result,
    }


@app.get("/storage/{day}/{filename}", tags=["storage"], summary="View or download a specific output file")
def storage_download(day: str, filename: str, download: bool = False):
    """
    Serves a file from `DOWNLOAD_PATH/<day>/<filename>`.

    **Rendering behaviour (default)**
    - `.log` — styled dark-theme monospace HTML with a download button
    - `.html` — rendered in the browser as-is
    - `.json` — returned with `application/json` content type
    - all others — downloaded as `application/octet-stream` attachment

    **`?download=true`** — forces any file type to download as an attachment.

    Path is validated against `DOWNLOAD_PATH` to prevent directory traversal.
    """
    resolved = _safe_resolve(day, filename)
    if not os.path.isfile(resolved):
        raise HTTPException(status_code=404, detail=f"{day}/{filename} not found")

    with open(resolved, "rb") as fh:
        raw = fh.read()

    if download:
        return StreamingResponse(
            io.BytesIO(raw),
            media_type = "application/octet-stream",
            headers    = {"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    if filename.endswith(".html"):
        return HTMLResponse(content=raw.decode("utf-8", errors="replace"))

    if filename.endswith(".log"):
        text      = raw.decode("utf-8", errors="replace")
        safe_name = html.escape(filename)
        safe_body = html.escape(text)
        return HTMLResponse(content=(
            f"<!DOCTYPE html><html><head><meta charset='UTF-8'/>"
            f"<title>{safe_name}</title></head>"
            f"<body style='background:#0f1117;color:#e2e8f0;font-family:monospace;padding:2rem'>"
            f"<div style='display:flex;align-items:center;gap:1rem;margin-bottom:1rem'>"
            f"<h2 style='color:#94a3b8;margin:0'>{safe_name}</h2>"
            f"<a href='/storage/{day}/{filename}?download=true' "
            f"style='font-size:.75rem;color:#60a5fa;text-decoration:none;border:1px solid #334155;"
            f"padding:.2rem .6rem;border-radius:4px'>⬇ download</a>"
            f"<a href='/storage/{day}' "
            f"style='font-size:.75rem;color:#94a3b8;text-decoration:none;border:1px solid #334155;"
            f"padding:.2rem .6rem;border-radius:4px'>← folder</a>"
            f"</div>"
            f"<pre style='font-size:.82rem;line-height:1.6;white-space:pre-wrap;"
            f"background:#1e293b;padding:1.25rem;border-radius:8px;border:1px solid #334155'>"
            f"{safe_body}</pre>"
            f"</body></html>"
        ))

    if filename.endswith(".json"):
        return StreamingResponse(
            io.BytesIO(raw),
            media_type = "application/json",
        )

    return StreamingResponse(
        io.BytesIO(raw),
        media_type = "application/octet-stream",
        headers    = {"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/storage/{day}", tags=["storage"], response_class=HTMLResponse, summary="Browse a date folder as an HTML page")
def storage_browse_folder(day: str):
    """
    Renders a dark-theme HTML directory listing for `DOWNLOAD_PATH/<day>/` with view and download
    links for each file. File type badges are color-coded: `.log` (blue), `.json` (purple), `.html` (green).

    Useful for quick visual inspection of a vectorize run's output without using the API directly.
    """
    safe_root = os.path.realpath(DOWNLOAD_PATH)
    day_path  = os.path.realpath(os.path.join(DOWNLOAD_PATH, day))
    if not day_path.startswith(safe_root + os.sep):
        raise HTTPException(status_code=400, detail="Invalid path")
    if not os.path.isdir(day_path):
        raise HTTPException(status_code=404, detail=f"Folder {day} not found")

    rows = ""
    for fname in sorted(os.listdir(day_path)):
        fpath = os.path.join(day_path, fname)
        if not os.path.isfile(fpath):
            continue
        size     = _fmt_size(os.path.getsize(fpath))
        modified = datetime.fromtimestamp(os.path.getmtime(fpath)).strftime("%Y-%m-%d %H:%M:%S")
        safe_f   = html.escape(fname)
        ext      = fname.rsplit(".", 1)[-1].lower() if "." in fname else ""
        badge    = {"log": "#60a5fa", "json": "#a78bfa", "html": "#34d399"}.get(ext, "#94a3b8")
        rows += (
            f"<tr>"
            f"<td><span style='background:{badge}22;color:{badge};border:1px solid {badge}44;"
            f"border-radius:4px;padding:.1rem .4rem;font-size:.7rem;font-weight:600'>{ext or '—'}</span></td>"
            f"<td style='font-family:monospace;font-size:.82rem'>{safe_f}</td>"
            f"<td style='color:#94a3b8'>{size}</td>"
            f"<td style='color:#64748b;font-size:.78rem'>{modified}</td>"
            f"<td style='white-space:nowrap'>"
            f"<a href='/storage/{day}/{safe_f}' style='color:#60a5fa;margin-right:.75rem;font-size:.78rem'>view</a>"
            f"<a href='/storage/{day}/{safe_f}?download=true' style='color:#a78bfa;margin-right:.75rem;font-size:.78rem'>download</a>"
            f"</td>"
            f"</tr>"
        )

    return HTMLResponse(content=(
        f"<!DOCTYPE html><html><head><meta charset='UTF-8'/>"
        f"<title>/tmp/{html.escape(day)}</title>"
        f"<style>*{{box-sizing:border-box;margin:0;padding:0}}"
        f"body{{font-family:'Segoe UI',system-ui,sans-serif;background:#0f1117;color:#e2e8f0;padding:2rem}}"
        f"h1{{font-size:1.2rem;color:#f8fafc;margin-bottom:.25rem}}"
        f"p{{color:#64748b;font-size:.82rem;margin-bottom:1.5rem}}"
        f"table{{width:100%;border-collapse:collapse;font-size:.85rem}}"
        f"th{{padding:.5rem .75rem;text-align:left;color:#64748b;font-size:.72rem;font-weight:600;"
        f"text-transform:uppercase;letter-spacing:.06em;border-bottom:1px solid #334155;background:#0f1117}}"
        f"td{{padding:.55rem .75rem;border-bottom:1px solid #1e293b;vertical-align:middle}}"
        f"tr:hover td{{background:#1e293b55}}"
        f"a{{color:#60a5fa;text-decoration:none}}a:hover{{text-decoration:underline}}"
        f"</style></head>"
        f"<body>"
        f"<h1>/tmp/{html.escape(day)}</h1>"
        f"<p><a href='/storage' style='color:#94a3b8'>← all folders</a></p>"
        f"<table><thead><tr><th>type</th><th>filename</th><th>size</th><th>modified</th><th>actions</th></tr></thead>"
        f"<tbody>{rows}</tbody></table>"
        f"</body></html>"
    ))


@app.delete("/storage/{day}/{filename}", tags=["storage"], summary="Delete a single output file", dependencies=[Depends(require_destructive)])
def storage_delete(day: str, filename: str):
    """
    Permanently deletes `DOWNLOAD_PATH/<day>/<filename>` from the local filesystem.

    Returns `{deleted, size_bytes}` on success. Raises **404** if the file does not exist.
    Path is validated to prevent directory traversal outside `DOWNLOAD_PATH`.
    """
    resolved = _safe_resolve(day, filename)
    if not os.path.isfile(resolved):
        raise HTTPException(status_code=404, detail=f"{day}/{filename} not found")
    size = os.path.getsize(resolved)
    os.remove(resolved)
    log.info(f"Deleted storage file: {day}/{filename} ({_fmt_size(size)})")
    return {"deleted": f"{day}/{filename}", "size_bytes": size}


@app.delete("/storage/{day}", tags=["storage"], summary="Delete an entire date folder and all its files", dependencies=[Depends(require_destructive)])
def storage_delete_folder(day: str):
    """
    Deletes all files inside `DOWNLOAD_PATH/<day>/` and then removes the directory itself.

    Returns `{deleted_folder, deleted_files, count}`. Raises **404** if the folder does not exist.
    This is a destructive operation — use it to free `/tmp` space after confirming outputs are
    backed up to S3 via `POST /s3/upload-storage-folder/{day}`.
    """
    safe_root = os.path.realpath(DOWNLOAD_PATH)
    day_path  = os.path.realpath(os.path.join(DOWNLOAD_PATH, day))
    if not day_path.startswith(safe_root + os.sep):
        raise HTTPException(status_code=400, detail="Invalid path")
    if not os.path.isdir(day_path):
        raise HTTPException(status_code=404, detail=f"Folder {day} not found")

    deleted = []
    for fname in os.listdir(day_path):
        fpath = os.path.join(day_path, fname)
        if os.path.isfile(fpath):
            os.remove(fpath)
            deleted.append(fname)
    os.rmdir(day_path)
    log.info(f"Deleted storage folder: {day} ({len(deleted)} files)")
    return {"deleted_folder": day, "deleted_files": deleted, "count": len(deleted)}


@app.get("/files", response_model=FilesResponse, dependencies=[Depends(require_connection)], tags=["files"], summary="List all SharePoint root files and sync metadata to HANA")
def list_files():
    """
    Fetches all files from the SharePoint root drive using the **Graph delta API**
    (`drives/{drive_id}/root/delta`) and upserts their metadata into `BRAIN_SP_FILES`.

    **What it does**
    1. Calls the Graph delta endpoint (paginates automatically via `@odata.nextLink`)
    2. Flattens each item into a metadata dict (id, name, type, extension, size, dates, URLs, shared scope)
    3. Upserts file records into `BRAIN_SP_FILES` — chunks and vectors are **not** created here
    4. Returns the full file list with the HANA upsert result

    **Use this when** you want a fast metadata sync without downloading or vectorizing files.
    Use `GET /vectorize` or `POST /vectorize` to also extract text and build vector embeddings.

    Requires an active SharePoint connection (returns 503 if not ready).
    """
    clog = _call_logger("files")
    clog.info("GET /files called")
    try:
        result = _get_root_files()
        clog.info(f"Fetched {len(result)} files from SharePoint")
    except requests.HTTPError as e:
        clog.error(f"HTTP error: {e.response.status_code}", exc_info=True)
        raise HTTPException(status_code=e.response.status_code, detail=str(e))
    except Exception as e:
        clog.error(f"Error: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))

    try:
        db_result = db.ingest([], [], result)
        clog.info(f"DB ingest complete — {db_result}")
    except Exception as e:
        clog.error(f"DB ingest failed: {e}")
        raise HTTPException(status_code=500, detail=f"DB ingest failed: {e}")

    return FilesResponse(count=len(result), files=result, db_ingest=db_result)


@app.get("/files/versions", dependencies=[Depends(require_connection)], tags=["files"], summary="List all root files with full version history")
def list_files_with_versions(
    delay: float = Query(1.0, ge=0.0, le=10.0, description="Seconds to sleep between per-file version requests to avoid Graph API throttling (429). Range: 0–10 s."),
):
    """
    Same as `GET /files` but also fetches the complete version history for every file.

    For each file, an additional Graph API call is made to
    `drives/{drive_id}/items/{id}/versions`. This can be slow for large libraries —
    use the `delay` parameter to throttle requests and avoid 429 rate limiting.

    **Response** — array of file dicts, each extended with:
    - `version_count` — total number of versions in SharePoint
    - `versions_list` — array of `{version_id, fileSize, size, lastModifiedDateTime, lastmodified_by, download_url}`

    **Note:** No HANA ingest is performed by this endpoint. Use `/pipeline` or `/vectorize` to
    ingest versions into `BRAIN_SP_FILES_VERSIONS`.
    """
    clog = _call_logger("files_versions")
    clog.info(f"GET /files/versions called — delay={delay}")
    try:
        result = _get_root_files_with_versions(delay=delay)
        clog.info(f"Returned {len(result)} files with versions")
        return result
    except requests.HTTPError as e:
        clog.error(f"HTTP error: {e.response.status_code}", exc_info=True)
        raise HTTPException(status_code=e.response.status_code, detail=str(e))
    except Exception as e:
        clog.error(f"Error: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/files/all", dependencies=[Depends(require_connection)], tags=["files"], summary="List files from a specific folder path (or root)")
def list_all_files(
    path:      str  = Query("",   description="SharePoint-relative folder path from the drive root, e.g. '/Documents/ProjectX'. Omit or leave empty to list the entire root."),
    use_delta: bool = Query(True, description="When true (default), uses the :/delta endpoint — recursive and change-tracked. When false, uses :/children — direct children only, no recursion."),
):
    """
    Flexible file listing endpoint. Lists files from a specified SharePoint folder path,
    or from the root if `path` is omitted.

    **`use_delta=true` (default)**
    Uses `drives/{drive_id}/root:{path}:/delta` — returns all files recursively under the path,
    including files in sub-folders. Paginates via `@odata.nextLink`.

    **`use_delta=false`**
    Uses `drives/{drive_id}/root:{path}` — returns only direct children (non-recursive).
    Faster for shallow folders but misses nested content.

    **No HANA ingest** — this is a read-only listing. Pass the result to `POST /vectorize`
    in the `items` field to vectorize a specific folder's contents.

    Example: `GET /files/all?path=/Documents/Reports&use_delta=true`
    """
    clog = _call_logger("files_all")
    clog.info(f"GET /files/all called — path='{path}' use_delta={use_delta}")
    try:
        result = _get_root_files() if not path else _get_all_files(path, use_delta=use_delta)
        clog.info(f"Returned {len(result)} files")
        return result
    except requests.HTTPError as e:
        clog.error(f"HTTP error: {e.response.status_code}", exc_info=True)
        raise HTTPException(status_code=e.response.status_code, detail=str(e))
    except Exception as e:
        clog.error(f"Error: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/files/search", dependencies=[Depends(require_connection)], tags=["files"], summary="Search SharePoint files by filename")
def search_files(
    name: str = Query(..., description="Filename or partial filename to search for (case-insensitive). Passed directly to the Graph search API: drives/{drive_id}/root/search(q='{name}')."),
):
    """
    Searches for files in the SharePoint drive whose name contains the given query string.

    Uses the Graph API search endpoint (`drives/{drive_id}/root/search(q='{name}')`) and
    additionally filters client-side to ensure the query string appears in the filename.
    Folders are excluded from results.

    **Response fields**
    - `query` — the search string that was used
    - `count` — number of matching files
    - `items` — array of flattened file metadata dicts (same shape as `/files`)

    **No HANA ingest** — this is a read-only search. Pass results to `POST /vectorize` to vectorize
    individual found files.

    Example: `GET /files/search?name=Q1+Report`
    """
    clog = _call_logger("files_search")
    clog.info(f"GET /files/search called — name='{name}'")
    try:
        results = _search_files_by_name(name)
        clog.info(f"Found {len(results)} matches for '{name}'")
        return {"query": name, "count": len(results), "items": results}
    except requests.HTTPError as e:
        status = e.response.status_code if e.response is not None else 500
        clog.error(f"HTTP error: {status}")
        raise HTTPException(status_code=status, detail=str(e))
    except Exception as e:
        clog.error(f"Search error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/files/hash/{item_id}", dependencies=[Depends(require_connection)], tags=["files"], summary="Get SharePoint file hash by item ID")
def get_file_hash(item_id: str):
    """
    Fetches a single DriveItem by ID from SharePoint and returns its `quickXorHash`
    (SharePoint's native file fingerprint). Use this to check whether two files with
    the same name are byte-identical without downloading them.
    """
    clog = _call_logger("files_hash")
    try:
        item = _get_item_by_id(item_id)
        url  = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
        raw  = requests.get(url, headers=_headers(), timeout=15).json()
        qxh  = raw.get("file", {}).get("hashes", {}).get("quickXorHash")
        sha1 = raw.get("file", {}).get("hashes", {}).get("sha1Hash")
        return {
            "item_id":        item_id,
            "name":           item.get("name"),
            "size":           item.get("size"),
            "last_modified":  item.get("last_modified_date"),
            "quick_xor_hash": qxh,
            "sha1_hash":      sha1,
        }
    except Exception as e:
        clog.error(f"Hash fetch failed for {item_id}: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get(
    "/vectorize",
    response_model = PipelineResponse,
    dependencies   = [Depends(require_connection)],
    tags           = ["rag"],
    summary        = "Full vectorize pipeline — fetch all root files, extract, embed, upsert (GET)",
)
def get_vectorize(
    save_dir: str   | None = Query(None, description="Absolute local path for output files. Defaults to DOWNLOAD_PATH/<YYYYMMDD>/."),
    delay:    float        = Query(1.0,  ge=0.0, le=10.0, description="Seconds between per-file version-fetch requests. Range: 0–10 s."),
):
    """
    End-to-end vectorization pipeline for **all files** in the SharePoint root drive.

    **Pipeline steps**
    1. Fetch all root files + version history via the Graph delta API (`delay` throttles version calls)
    2. Compare `lastModifiedDateTime` against HANA — skip files already up-to-date
    3. For each new/modified file: download → extract text → semantic chunk (512-token windows) → embed (all-MiniLM-L6-v2, 384-dim) → upsert to HANA
    4. Upsert version records to `BRAIN_SP_FILES_VERSIONS`
    5. Write output files to `save_dir` (chunks, vectors, raw-chunks, root JSON, versions JSON)
    6. Generate an HTML analysis report and upload it + the call log to S3

    **Change detection** — only files whose `lastModifiedDateTime` has changed since the last
    run are re-processed. The `skipped_unchanged` field shows how many were skipped.

    **Supported formats** — docx, xlsx, pdf, pptx, csv, json, xml, html, txt, py, zip (recursive), and more.

    For a specific folder use `GET /vectorize/folder/{path}`.
    For a POST body with an explicit item list use `POST /vectorize`.
    """
    clog = _call_logger("vectorize")
    clog.info(f"GET /vectorize called — delay={delay}, save_dir={save_dir}")
    log.info("GET /vectorize — fetching root files with versions")
    try:
        root_files = _get_root_files_with_versions(delay=delay)
    except Exception as e:
        clog.error(f"Failed to fetch root files: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to fetch root files: {e}")

    if not root_files:
        clog.warning("No files found in SharePoint root")
        raise HTTPException(status_code=404, detail="No files found in SharePoint root")

    root_files, versions_data = _split_versions(root_files)
    clog.info(f"Fetched {len(root_files)} root files ({len(versions_data)} version records)")
    # try:  # disabled — BRAIN_FILE_REGISTRY
    #     db.ingest_file_registry(root_files, source="sharepoint")
    # except Exception as _e:
    #     clog.warning(f"file_registry ingest skipped: {_e}")
    to_process, skipped_unchanged = _filter_stale_items(root_files, logger=clog)
    clog.info(f"Starting vectorization — {len(to_process)} files to process, {skipped_unchanged} unchanged")
    if not to_process:
        clog.info("All files already up-to-date — nothing to vectorize")
        return PipelineResponse(
            count_root_files  = len(root_files),
            count_chunks      = 0,
            count_files       = 0,
            skipped_unchanged = skipped_unchanged,
            saved_to          = [],
            db_ingest         = {"message": "all files up-to-date"},
        )
    try:
        n_chunks, n_files, path_chunks, path_raw, path_vectors, path_html, path_log, ingest_results = _download_and_vectorize(to_process, logger=clog)
    except Exception as e:
        clog.error(f"Vectorization failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Vectorization failed: {e}")

    ts            = datetime.now().strftime("%Y%m%d_%H%M%S")
    save_dir      = save_dir or os.path.dirname(path_chunks)
    path_root     = os.path.join(save_dir, f"sp_root_data_{ts}.json")
    path_versions = os.path.join(save_dir, f"sp_files_versions_{ts}.json")
    _write_to_json_array_indent(root_files,    path=path_root)
    _write_to_json_array_indent(versions_data, path=path_versions)

    log.info(f"GET /vectorize complete — {len(to_process)} files processed, {n_chunks} chunks, {skipped_unchanged} skipped → {save_dir}")
    clog.info(f"Complete — {len(to_process)} processed, {n_chunks} chunks, {skipped_unchanged} unchanged")

    db_result = {"per_file_ingest": ingest_results}

    return PipelineResponse(
        count_root_files  = len(root_files),
        count_chunks      = n_chunks,
        count_files       = n_files,
        skipped_unchanged = skipped_unchanged,
        saved_to          = [p for p in [path_chunks, path_raw, path_vectors, path_root, path_versions, path_html, path_log] if p],
        db_ingest         = db_result,
    )


@app.get(
    "/vectorize/folder/{folder_path:path}",
    response_model = PipelineResponse,
    dependencies   = [Depends(require_connection)],
    tags           = ["rag"],
    summary        = "Vectorize all files under a specific SharePoint subfolder",
)
def vectorize_folder(
    folder_path: str,
    delay:     float = Query(1.0,  ge=0.0, le=10.0, description="Seconds between per-file version-fetch requests to avoid 429 throttling. Range: 0–10 s."),
    use_delta: bool  = Query(True, description="When true (default), uses :/delta — recursive and change-tracked. When false, uses :/children — direct children only."),
):
    """
    Same pipeline as `GET /vectorize` but scoped to a single SharePoint subfolder.

    **Graph API call**
    - `use_delta=true` (default): `drives/{drive_id}/root:/{folder_path}:/delta` — recursive, all sub-folders
    - `use_delta=false`: `drives/{drive_id}/root:/{folder_path}` — direct children only

    **Pipeline steps**
    1. Fetch all files under the folder path (paginated)
    2. Fetch version history per file (`delay` throttles these calls)
    3. Skip files unchanged since last HANA upsert
    4. For each changed file: download → extract → chunk → embed → upsert to HANA
    5. Upsert version records to `BRAIN_SP_FILES_VERSIONS`
    6. Generate HTML report + upload artifacts to S3

    **Example:** `GET /vectorize/folder/Documents/ProjectX`

    The `folder_path` is URL-encoded automatically — spaces become `%20`. Slashes in the path are passed through as path segments.
    """
    clog = _call_logger("vectorize_folder")
    clog.info(f"GET /vectorize/folder/{folder_path} called — delay={delay} use_delta={use_delta}")

    src = f"/{folder_path.lstrip('/')}"
    try:
        items = _get_all_files(src, use_delta=use_delta)
    except Exception as e:
        clog.error(f"Failed to fetch folder '{src}': {e}")
        raise HTTPException(status_code=500, detail=f"Failed to fetch folder '{src}': {e}")

    if not items:
        clog.warning(f"No files found in folder '{src}'")
        raise HTTPException(status_code=404, detail=f"No files found in '{src}'")

    clog.info(f"Fetched {len(items)} files from '{src}' — fetching versions")
    versions_data = []
    for _item in items:
        _vlist = _get_item_versions(_item["id"])
        if _vlist:
            versions_data.extend(_vlist)
        if delay:
            time.sleep(delay)
    clog.info(f"Fetched {len(versions_data)} version records")

    # try:  # disabled — BRAIN_FILE_REGISTRY
    #     db.ingest_file_registry(items, source="sharepoint")
    # except Exception as _e:
    #     clog.warning(f"file_registry ingest skipped: {_e}")
    to_process, skipped_unchanged = _filter_stale_items(items, logger=clog)
    clog.info(f"Starting vectorization — {len(to_process)} files to process, {skipped_unchanged} unchanged")

    db_result = {}

    if not to_process:
        clog.info("All files already up-to-date — nothing to vectorize")
        db_result["message"] = "all files up-to-date"
        return PipelineResponse(
            count_root_files  = len(items),
            count_chunks      = 0,
            count_files       = 0,
            skipped_unchanged = skipped_unchanged,
            saved_to          = [],
            db_ingest         = db_result,
        )
    try:
        n_chunks, n_files, path_chunks, path_raw, path_vectors, path_html, path_log, ingest_results = _download_and_vectorize(to_process, logger=clog)
    except Exception as e:
        clog.error(f"Vectorization failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Vectorization failed: {e}")

    db_result["per_file_ingest"] = ingest_results
    clog.info(f"Complete — {len(to_process)} processed, {n_chunks} chunks, {skipped_unchanged} unchanged from '{src}'")
    return PipelineResponse(
        count_root_files  = len(items),
        count_chunks      = n_chunks,
        count_files       = n_files,
        skipped_unchanged = skipped_unchanged,
        saved_to          = [p for p in [path_chunks, path_raw, path_vectors, path_html, path_log] if p],
        db_ingest         = db_result,
    )


@app.post(
    "/vectorize/s3",
    response_model = VectorizeResponse,
    tags           = ["rag"],
    summary        = "Vectorize all files under an S3 folder prefix",
)
def vectorize_s3(req: S3VectorizeRequest, skip_ingest: bool = False):
    """
    Picks up every file under the given S3 **prefix**, downloads each one via the
    configured Object Store binding, extracts text, chunks, embeds with
    `all-MiniLM-L6-v2`, and upserts chunks + vectors into HANA.

    **Change detection** — unless `force=true`, files whose `last_modified`
    timestamp already matches the value stored in HANA are skipped.

    **Supported file types** — same as the SharePoint pipeline:
    docx, doc, xlsx, xls, pdf, pptx, txt, csv, json, md, rtf, msg, html, py, js, ts, yaml, sh

    **Body example**
    ```json
    {
      "prefix": "BTP-Catalyst/knowledge-base/",
      "force": false,
      "extensions": ["pdf", "docx", "xlsx"]
    }
    ```
    `extensions` filters to only those types; omit or pass `[]` for all supported types.
    """
    import s3 as s3_client

    clog = _call_logger("vectorize_s3")
    clog.info(f"POST /vectorize/s3 — prefix={req.prefix!r} force={req.force} extensions={req.extensions}")

    # ── 1. List all objects under the prefix ─────────────────────────────────
    SUPPORTED = {
        "docx","doc","xlsx","xlsm","xls","pdf","pptx","pptm",
        "txt","csv","json","md","rtf","msg","html","htm","py","js","ts","yaml","yml","sh",
    }
    ext_filter = {e.lower().lstrip(".") for e in req.extensions} if req.extensions else None

    try:
        raw_objects = s3_client.list_objects(prefix=req.prefix)
    except Exception as e:
        clog.error(f"S3 list failed: {e}")
        raise HTTPException(status_code=500, detail=f"S3 list_objects failed: {e}")

    # Filter to supported / requested extensions
    objects = []
    for obj in raw_objects:
        key = obj["key"]
        if key.endswith("/"):          # skip folder markers
            continue
        filename = key.split("/")[-1]
        ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
        if ext not in SUPPORTED:
            continue
        if ext_filter and ext not in ext_filter:
            continue
        objects.append(obj)

    clog.info(f"Found {len(raw_objects)} total objects, {len(objects)} after extension filter")

    if not objects:
        raise HTTPException(
            status_code=404,
            detail=f"No supported files found under prefix '{req.prefix}'. "
                   f"Supported types: {sorted(SUPPORTED)}",
        )

    # ── 2. Change detection ───────────────────────────────────────────────────
    if req.force:
        to_process       = objects
        skipped_unchanged = 0
        clog.info("force=true — processing all objects regardless of HANA state")
    else:
        items_for_filter = [_s3_obj_to_item(o) for o in objects]
        filtered, skipped_unchanged = _filter_stale_items(items_for_filter, logger=clog)
        # Remap back to original object dicts using etag (item_id) → key mapping
        processed_etags = {item["id"] for item in filtered}
        to_process      = [o for o in objects if o.get("etag", o["key"]) in processed_etags]

    clog.info(f"{len(to_process)} to vectorize, {skipped_unchanged} unchanged (skipped)")

    if not to_process:
        clog.info("All S3 files already up-to-date in HANA")
        return VectorizeResponse(
            count_chunks      = 0,
            count_files       = 0,
            skipped_unchanged = skipped_unchanged,
            saved_to          = [],
            db_ingest         = {"message": "all files up-to-date"},
        )

    # ── 3. Download → extract → chunk → embed → HANA ─────────────────────────
    try:
        n_chunks, n_files, path_chunks, path_raw, path_vectors, path_root, path_html, path_log, ingest_results = \
            _download_and_vectorize_s3(to_process, logger=clog, skip_ingest=skip_ingest)
    except Exception as e:
        clog.error(f"S3 vectorization failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"S3 vectorization failed: {e}")

    clog.info(f"Complete — {n_files} files, {n_chunks} chunks, {skipped_unchanged} unchanged")
    return VectorizeResponse(
        count_chunks      = n_chunks,
        count_files       = n_files,
        skipped_unchanged = skipped_unchanged,
        saved_to          = [p for p in [path_chunks, path_raw, path_vectors, path_root, path_html, path_log] if p],
        db_ingest         = {"per_file_ingest": ingest_results},
    )


@app.get(
    "/vectorize/s3/folder/{folder_path:path}",
    response_model = VectorizeResponse,
    tags           = ["rag"],
    summary        = "Vectorize all files under an S3 folder (GET, path-based)",
)
def vectorize_s3_folder(
    folder_path: str,
    force: bool = Query(False, description="Re-vectorize all files even if unchanged in HANA"),
    extensions: str = Query("", description="Comma-separated extension whitelist, e.g. 'pdf,docx'. Empty = all supported types."),
):
    """
    Same pipeline as `POST /vectorize/s3` but accepts the folder as a URL path segment.

    **Example:** `GET /vectorize/s3/folder/BTP-Catalyst/knowledge-base`

    Trailing slash is optional — the prefix passed to S3 will always end with `/`.

    Query params:
    - `force` — skip change detection, re-vectorize everything
    - `extensions` — comma-separated list, e.g. `pdf,docx,xlsx`
    """
    ext_list = [e.strip().lower().lstrip(".") for e in extensions.split(",") if e.strip()] if extensions else []
    prefix   = folder_path.rstrip("/") + "/"
    req      = S3VectorizeRequest(prefix=prefix, force=force, extensions=ext_list)
    return vectorize_s3(req, skip_ingest=False)


@app.get(
    "/vectorize/s3/folder/{folder_path:path}/download",
    tags    = ["rag"],
    summary = "Vectorize an S3 folder and download the output as a zip (no HANA ingest)",
)
def vectorize_s3_folder_download(
    folder_path: str,
    force:       bool = Query(False, description="Re-vectorize all files even if unchanged in HANA"),
    extensions:  str  = Query("",    description="Comma-separated extension whitelist, e.g. 'pdf,docx'. Empty = all supported types."),
):
    """
    Same pipeline as `GET /vectorize/s3/folder/{folder_path}` but instead of returning JSON,
    streams all output JSONL files (chunks, vectors, raw-chunk summary) packaged as a `.zip`
    attachment — ready to save directly to your PC.

    HANA ingest is **disabled** — this endpoint is for extracting vectorized data locally.

    **Returns:** `application/zip` — filename pattern `s3_vectorize_<prefix>_<timestamp>.zip`

    Contents:
    - `s3_chunks_<ts>.jsonl`   — text chunk records
    - `s3_vectors_<ts>.jsonl`  — embedding vectors (384-dim)
    - `s3_files_chunks_<ts>.json` — raw binary chunks (base64-encoded, for rebuild)
    """
    import zipfile

    clog     = _call_logger("vectorize_s3_folder_download")
    ext_list = [e.strip().lower().lstrip(".") for e in extensions.split(",") if e.strip()] if extensions else []
    prefix   = folder_path.rstrip("/") + "/"
    clog.info(f"GET /vectorize/s3/folder/{folder_path}/download — prefix={prefix!r} force={force} extensions={ext_list}")

    import s3 as s3_client

    SUPPORTED = {
        "docx","doc","xlsx","xlsm","xls","pdf","pptx","pptm",
        "txt","csv","json","md","rtf","msg","html","htm","py","js","ts","yaml","yml","sh",
    }
    ext_filter = {e.lower().lstrip(".") for e in ext_list} if ext_list else None

    try:
        raw_objects = s3_client.list_objects(prefix=prefix)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"S3 list_objects failed: {e}")

    objects = []
    for obj in raw_objects:
        key = obj["key"]
        if key.endswith("/"):
            continue
        filename = key.split("/")[-1]
        ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
        if ext not in SUPPORTED:
            continue
        if ext_filter and ext not in ext_filter:
            continue
        objects.append(obj)

    if not objects:
        raise HTTPException(
            status_code=404,
            detail=f"No supported files found under prefix '{prefix}'. Supported: {sorted(SUPPORTED)}",
        )

    if not force:
        items_for_filter = [_s3_obj_to_item(o) for o in objects]
        filtered, _ = _filter_stale_items(items_for_filter, logger=clog)
        processed_etags = {item["id"] for item in filtered}
        objects = [o for o in objects if o.get("etag", o["key"]) in processed_etags]

    if not objects:
        raise HTTPException(status_code=200, detail="All S3 files already up-to-date — nothing to vectorize")

    try:
        n_chunks, n_files, path_chunks, path_raw, path_vectors, path_root, path_html, path_log, _ = \
            _download_and_vectorize_s3(objects, logger=clog, skip_ingest=True)
    except Exception as e:
        clog.error(f"Vectorization failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Vectorization failed: {e}")

    # Package output JSONLs into an in-memory zip
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in [path_chunks, path_raw, path_vectors, path_root]:
            if path and os.path.exists(path):
                zf.write(path, arcname=os.path.basename(path))
    zip_buf.seek(0)

    safe_prefix = prefix.strip("/").replace("/", "_")[:40]
    ts          = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_name    = f"s3_vectorize_{safe_prefix}_{ts}.zip"
    clog.info(f"Streaming zip — {n_files} files, {n_chunks} chunks → {zip_name}")

    return StreamingResponse(
        zip_buf,
        media_type = "application/zip",
        headers    = {"Content-Disposition": f'attachment; filename="{zip_name}"'},
    )


@app.post(
    "/vectorize/s3/item",
    response_model = S3VectorizeItemResponse,
    tags           = ["rag"],
    summary        = "Vectorize a single S3 object by its key",
)
def vectorize_s3_item(req: S3VectorizeItemRequest):
    """
    Fetches and vectorizes a single file from S3 by its full object key.

    **Pipeline steps**
    1. `s3.get_metadata(key)` — verify the object exists and get `last_modified`
    2. Change detection against HANA — return immediately with `skipped_unchanged=1` if up-to-date
    3. `s3.download_bytes(key)` → extract → chunk → embed → upsert to HANA
    4. Write JSONL output files, generate HTML report, upload artifacts to S3

    **Body example**
    ```json
    { "key": "BTP-Catalyst/knowledge-base/report.pdf", "force": false }
    ```
    """
    import s3 as s3_client

    clog     = _call_logger("vectorize_s3_item")
    key      = req.key.lstrip("/")
    filename = key.split("/")[-1] if "/" in key else key
    ext      = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    clog.info(f"POST /vectorize/s3/item — key={key!r} force={req.force}")

    SUPPORTED = {
        "docx","doc","xlsx","xlsm","xls","pdf","pptx","pptm",
        "txt","csv","json","md","rtf","msg","html","htm","py","js","ts","yaml","yml","sh",
    }
    if ext not in SUPPORTED:
        raise HTTPException(
            status_code=422,
            detail=f"File extension '.{ext}' is not supported for text extraction. Supported: {sorted(SUPPORTED)}",
        )

    # ── Verify object exists & get metadata ───────────────────────────────────
    try:
        meta = s3_client.get_metadata(key)
    except Exception as e:
        raise HTTPException(status_code=404, detail=f"S3 object not found or inaccessible: {key} — {e}")

    obj = {
        "key":           key,
        "size":          meta.get("content_length", 0),
        "last_modified": meta.get("last_modified", ""),
        "etag":          meta.get("etag", ""),
        "storage_class": "STANDARD",
    }
    item = _s3_obj_to_item(obj)

    # ── Change detection ──────────────────────────────────────────────────────
    if not req.force:
        filtered, skipped = _filter_stale_items([item], logger=clog)
        if skipped:
            clog.info(f"Skipping {key!r} — already up-to-date in HANA")
            return S3VectorizeItemResponse(
                key               = key,
                filename          = filename,
                count_chunks      = 0,
                skipped_unchanged = 1,
                saved_to          = [],
                db_ingest         = {"message": "already up-to-date"},
            )

    # ── Download → extract → chunk → embed → HANA ────────────────────────────
    try:
        n_chunks, n_files, path_chunks, path_raw, path_vectors, path_root, path_html, path_log, ingest_results = \
            _download_and_vectorize_s3([obj], logger=clog)
    except Exception as e:
        clog.error(f"S3 item vectorization failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"S3 item vectorization failed: {e}")

    clog.info(f"Complete — key={key!r} chunks={n_chunks}")
    return S3VectorizeItemResponse(
        key               = key,
        filename          = filename,
        count_chunks      = n_chunks,
        skipped_unchanged = 0,
        saved_to          = [p for p in [path_chunks, path_raw, path_vectors, path_root, path_html, path_log] if p],
        db_ingest         = {"per_file_ingest": ingest_results},
    )


@app.post(
    "/vectorize",
    response_model = VectorizeResponse,
    dependencies   = [Depends(require_connection)],
    tags           = ["rag"],
    summary        = "Vectorize an explicit list of SharePoint items (POST body)",
)
def vectorize(req: VectorizeRequest):
    """
    Vectorizes a caller-supplied list of SharePoint item dicts, or auto-fetches all root files
    if `items` is omitted.

    **When `items` is provided**
    Pass the array returned by `GET /files`, `GET /files/all`, or `GET /files/search`.
    Only the items in the list are processed — useful for targeted re-indexing.

    **When `items` is omitted**
    Behaves identically to `GET /vectorize` — fetches the full root delta, then applies change
    detection and vectorizes only new/modified files.

    **Pipeline steps** (same as GET /vectorize)
    1. Change detection against HANA
    2. Download → extract text → semantic chunk → embed → upsert to HANA
    3. Optionally fetch + upsert version history if versions were included in the input
    4. Write output JSONL files, generate HTML report, upload artifacts to S3

    **Body example**
    ```json
    {
      "items": [ { "id": "...", "name": "report.pdf", "BaseURL": "...", ... } ],
      "save_dir": null
    }
    ```
    """
    clog = _call_logger("vectorize_post")
    clog.info(f"POST /vectorize called — items={len(req.items or [])}, save_dir={req.save_dir}")
    items         = req.items or []
    versions_data = []

    if not items:
        log.info("No items supplied — fetching root files with versions")
        clog.info("No items supplied — fetching root files with versions")
        try:
            fetched = _get_root_files_with_versions()
            items, versions_data = _split_versions(fetched)
        except Exception as e:
            clog.error(f"Failed to fetch root files: {e}")
            raise HTTPException(status_code=500, detail=f"Failed to fetch root files: {e}")

    if not items:
        clog.warning("No files found to vectorize")
        raise HTTPException(status_code=404, detail="No files found to vectorize")

    # try:  # disabled — BRAIN_FILE_REGISTRY
    #     db.ingest_file_registry(items, source="sharepoint")
    # except Exception as _e:
    #     clog.warning(f"file_registry ingest skipped: {_e}")
    to_process, skipped_unchanged = _filter_stale_items(items, logger=clog)
    clog.info(f"Starting vectorization — {len(to_process)} to process, {skipped_unchanged} unchanged")
    if not to_process:
        clog.info("All files already up-to-date — nothing to vectorize")
        return VectorizeResponse(
            count_chunks      = 0,
            count_files       = 0,
            skipped_unchanged = skipped_unchanged,
            saved_to          = [],
            db_ingest         = {"message": "all files up-to-date"},
        )
    try:
        n_chunks, n_files, path_chunks, path_raw, path_vectors, path_html, path_log, ingest_results = _download_and_vectorize(to_process, logger=clog)
    except Exception as e:
        clog.error(f"Vectorization failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Vectorization failed: {e}")

    saved_to  = [p for p in [path_chunks, path_raw, path_vectors, path_html, path_log] if p]
    db_result = {"per_file_ingest": ingest_results}

    if versions_data:
        path_versions = os.path.join(os.path.dirname(path_chunks), f"sp_files_versions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        _write_to_json_array_indent(versions_data, path=path_versions)
        saved_to.append(path_versions)


    log.info(f"Vectorize complete — {n_chunks} chunks, {n_files} processed, {skipped_unchanged} skipped → {path_chunks}")
    clog.info(f"Complete — {n_chunks} chunks, {n_files} processed, {skipped_unchanged} unchanged")

    return VectorizeResponse(
        count_chunks      = n_chunks,
        count_files       = n_files,
        skipped_unchanged = skipped_unchanged,
        saved_to          = saved_to,
        db_ingest         = db_result,
    )


@app.post(
    "/vectorize/item",
    response_model = VectorizeItemResponse,
    dependencies   = [Depends(require_connection)],
    tags           = ["rag"],
    summary        = "Fetch and vectorize a single SharePoint item by its item ID",
)
def vectorize_single_item(req: VectorizeItemRequest):
    """
    Fetches a single SharePoint file by its Graph API item ID, applies change detection,
    and vectorizes it if it is new or modified.

    **Pipeline steps**
    1. `GET drives/{drive_id}/items/{item_id}` — resolve file metadata
    2. `GET drives/{drive_id}/items/{item_id}/versions` — fetch version history
    3. Upsert version records to `BRAIN_SP_FILES_VERSIONS`
    4. Compare `lastModifiedDateTime` against HANA — return immediately with `skipped_unchanged=1` if up-to-date
    5. Download → extract text → semantic chunk → embed → upsert to HANA
    6. Write output JSONL files, generate HTML report, upload artifacts to S3

    **Body example**
    ```json
    { "item_id": "01ABCDEF1234567890ABCDEF" }
    ```

    Obtain the `item_id` from the `id` field in any `/files` or `/files/search` response.
    """
    clog = _call_logger("vectorize_item")
    clog.info(f"POST /vectorize/item called — item_id={req.item_id}")
    try:
        item = _get_item_by_id(req.item_id)
    except requests.HTTPError as e:
        status = e.response.status_code if e.response is not None else 500
        clog.error(f"Item lookup failed: {status}")
        raise HTTPException(status_code=status, detail=f"Item not found: {req.item_id}")
    except Exception as e:
        clog.error(f"Item lookup error: {e}")
        raise HTTPException(status_code=500, detail=str(e))

    clog.info(f"Fetched item: {item['name']} — fetching versions")
    versions_data = _get_item_versions(req.item_id)
    clog.info(f"Fetched {len(versions_data)} version records for {item['name']}")

    item_db_result = {}

    # try:  # disabled — BRAIN_FILE_REGISTRY
    #     db.ingest_file_registry([item], source="sharepoint")
    # except Exception as _e:
    #     clog.warning(f"file_registry ingest skipped: {_e}")
    to_process, skipped_unchanged = _filter_stale_items([item], logger=clog)
    if not to_process:
        clog.info(f"Item {item['name']} already up-to-date — skipping")
        item_db_result["message"] = "file already up-to-date"
        return VectorizeItemResponse(
            item_id           = req.item_id,
            filename          = item["name"],
            count_chunks      = 0,
            skipped_unchanged = 1,
            saved_to          = [],
            db_ingest         = item_db_result,
        )
    try:
        n_chunks, n_files, path_chunks, path_raw, path_vectors, path_html, path_log, ingest_results = _download_and_vectorize(to_process, logger=clog)
    except Exception as e:
        clog.error(f"Vectorization failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Vectorization failed: {e}")

    item_db_result["per_file_ingest"] = ingest_results
    return VectorizeItemResponse(
        item_id           = req.item_id,
        filename          = item["name"],
        count_chunks      = n_chunks,
        skipped_unchanged = skipped_unchanged,
        saved_to          = [p for p in [path_chunks, path_raw, path_vectors, path_html, path_log] if p],
        db_ingest         = item_db_result,
    )


@app.post(
    "/pipeline",
    response_model = PipelineResponse,
    dependencies   = [Depends(require_connection)],
    tags           = ["rag"],
    summary        = "Full pipeline — fetch root files with versions, vectorize, and save all outputs",
)
def pipeline(req: PipelineRequest = PipelineRequest()):
    """
    The most comprehensive pipeline endpoint. Fetches every root file **with full version history**
    and runs the complete vectorization flow.

    **Difference from `GET /vectorize`**
    - Always fetches version history for every file before vectorizing (slower, but `BRAIN_SP_FILES_VERSIONS` is guaranteed to be populated)
    - Saves `sp_root_data_<ts>.json` and `sp_files_versions_<ts>.json` in addition to chunk/vector/raw outputs
    - Ingest of `BRAIN_SP_FILES_VERSIONS` happens before vectorization (so versions are stored even if vectorize fails)

    **Pipeline steps**
    1. Fetch root files with version history (`_get_root_files_with_versions`, throttled by `delay`)
    2. Upsert version records to `BRAIN_SP_FILES_VERSIONS`
    3. Change detection — skip files unchanged in HANA
    4. Download → extract → chunk → embed → upsert to HANA
    5. Write all output files, generate HTML report, upload to S3

    **Body** — all fields optional (defaults shown):
    ```json
    { "save_dir": null, "delay": 1.0 }
    ```
    """
    clog = _call_logger("pipeline")
    clog.info(f"POST /pipeline called — delay={req.delay}, save_dir={req.save_dir}")
    log.info("Pipeline started — fetching root files with versions")
    try:
        root_files, versions_data = _split_versions(_get_root_files_with_versions(delay=req.delay))
    except Exception as e:
        clog.error(f"Failed to fetch root files: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to fetch root files: {e}")

    if not root_files:
        clog.warning("No files found in SharePoint root")
        raise HTTPException(status_code=404, detail="No files found in SharePoint root")

    clog.info(f"Fetched {len(root_files)} root files ({len(versions_data)} version records)")
    # try:  # disabled — BRAIN_FILE_REGISTRY
    #     db.ingest_file_registry(root_files, source="sharepoint")
    # except Exception as _e:
    #     clog.warning(f"file_registry ingest skipped: {_e}")
    to_process, skipped_unchanged = _filter_stale_items(root_files, logger=clog)
    clog.info(f"Starting vectorization — {len(to_process)} to process, {skipped_unchanged} unchanged")

    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    save_dir = req.save_dir or os.path.join(DOWNLOAD_PATH, ts[:8])
    os.makedirs(save_dir, exist_ok=True)
    path_root     = os.path.join(save_dir, f"sp_root_data_{ts}.json")
    path_versions = os.path.join(save_dir, f"sp_files_versions_{ts}.json")
    _write_to_json_array_indent(root_files,    path=path_root)
    _write_to_json_array_indent(versions_data, path=path_versions)

    db_result = {}

    if not to_process:
        clog.info("All files already up-to-date — nothing to vectorize")
        log.info(f"Pipeline complete — all {len(root_files)} files up-to-date, nothing vectorized")
        db_result["message"] = "all files up-to-date"
        return PipelineResponse(
            count_root_files  = len(root_files),
            count_chunks      = 0,
            count_files       = 0,
            skipped_unchanged = skipped_unchanged,
            saved_to          = [path_root, path_versions],
            db_ingest         = db_result,
        )
    try:
        n_chunks, n_files, path_chunks, path_raw, path_vectors, path_html, path_log, ingest_results = _download_and_vectorize(to_process, logger=clog)
    except Exception as e:
        clog.error(f"Vectorization failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Vectorization failed: {e}")

    log.info(f"Pipeline complete — {len(to_process)} processed, {n_chunks} chunks, {skipped_unchanged} skipped → {save_dir}")
    clog.info(f"Complete — {len(to_process)} processed, {n_chunks} chunks, {skipped_unchanged} unchanged")

    db_result["per_file_ingest"] = ingest_results

    return PipelineResponse(
        count_root_files  = len(root_files),
        count_chunks      = n_chunks,
        count_files       = n_files,
        skipped_unchanged = skipped_unchanged,
        saved_to          = [p for p in [path_chunks, path_raw, path_vectors, path_root, path_versions, path_html, path_log] if p],
        db_ingest         = db_result,
    )


@app.post("/download", dependencies=[Depends(require_connection)], tags=["files"], summary="Download a SharePoint file and stream it back")
def download_file(req: DownloadRequest):
    """
    Downloads a SharePoint file using the `BaseURL` from its metadata dict and streams the raw
    bytes back to the caller as an attachment.

    Optionally saves the file to `save_dir` and writes base64 chunk records to
    `files_list_chunk.json` in the same directory (used by `GET /rebuild` for offline replay).

    **Request body example**
    ```json
    {
      "item": { "id": "...", "name": "report.pdf", "BaseURL": "https://graph.microsoft.com/...", "size": "1.2 MB" },
      "save_dir": null
    }
    ```

    Uses streaming download with 8 KB chunks and a 300 s read timeout.
    Returns **404** if the download fails or the file has no content.
    """
    filename = req.item.get("name", "file")
    clog = _call_logger("download")
    clog.info(f"POST /download called — file={filename}, save_dir={req.save_dir}")
    content = _download_file_content(req.item, save_dir=req.save_dir)
    if not content:
        clog.error(f"File not found or download failed: {filename}")
        raise HTTPException(status_code=404, detail="File not found or download failed")
    clog.info(f"Downloaded {len(content) / 1024:.1f} KB — streaming back to caller")
    return StreamingResponse(
        io.BytesIO(content),
        media_type = "application/octet-stream",
        headers    = {"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/rebuild", tags=["files"], summary="Rebuild a file from stored chunk records (reads files_list_chunk.json)")
def rebuild_file_get(
    filename: str        = Query(...,  description="Display filename to rebuild, e.g. 'report.pdf'. Must match the 'filename' field in the chunk records."),
    save_dir: str | None = Query(None, description="Directory containing files_list_chunk.json. Defaults to DOWNLOAD_PATH/<today>/. Use the same save_dir that was passed to POST /download."),
):
    """
    Reconstructs a binary file from the base64 chunk records previously written by `POST /download`.

    Reads `files_list_chunk.json` (a JSONL file) from `save_dir`, filters records by `filename`,
    sorts them by `chunk_id`, base64-decodes each chunk, and reassembles the original bytes.
    The reconstructed file is streamed back as an attachment.

    **Use case:** re-serve a file that was already downloaded from SharePoint without making
    another Graph API call. Useful for inspection or re-processing without network access.

    Returns **404** if no chunks are found for the requested filename in the JSONL file.
    """
    clog = _call_logger("rebuild")
    clog.info(f"GET /rebuild called — filename={filename}, save_dir={save_dir}")
    active_dir  = datetime.now().strftime("%Y%m%d")
    resolved    = save_dir or os.path.join(DOWNLOAD_PATH, active_dir)
    chunk_path  = os.path.join(resolved, "files_list_chunk.json")

    chunks = _read_chunks_from_jsonl(chunk_path, filename)
    if not chunks:
        clog.error(f"No chunks found for '{filename}' in {chunk_path}")
        raise HTTPException(
            status_code = 404,
            detail      = f"No chunks found for '{filename}' in {chunk_path}",
        )

    clog.info(f"Found {len(chunks)} chunks — rebuilding")
    content = _rebuild_file(chunks, save_dir=save_dir)
    if not content:
        clog.error("Rebuild failed — check chunk data")
        raise HTTPException(status_code=400, detail="Rebuild failed — check chunk data")

    clog.info(f"Rebuilt {len(content) / 1024:.1f} KB — streaming back to caller")
    return StreamingResponse(
        io.BytesIO(content),
        media_type = "application/octet-stream",
        headers    = {"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/rebuild", tags=["files"], summary="Rebuild a file from a supplied chunk list (POST body)")
def rebuild_file(req: RebuildRequest):
    """
    Reconstructs a binary file from a list of base64-encoded chunk dicts supplied directly in
    the request body.

    Useful when you already have the chunk records in memory or fetched them from another source.
    Chunks are sorted by `chunk_id` before reassembly.

    **Request body example**
    ```json
    {
      "chunks": [
        { "chunk_id": 1, "chunk_data": "<base64>", "filename": "report.pdf" },
        { "chunk_id": 2, "chunk_data": "<base64>", "filename": "report.pdf" }
      ],
      "save_dir": null
    }
    ```

    Returns the reconstructed file as a streaming attachment. Optionally writes to `save_dir`.
    """
    clog = _call_logger("rebuild_post")
    clog.info(f"POST /rebuild called — chunks={len(req.chunks)}, save_dir={req.save_dir}")
    if not req.chunks:
        clog.error("chunks list is empty")
        raise HTTPException(status_code=422, detail="chunks list is empty")
    content = _rebuild_file(req.chunks, save_dir=req.save_dir)
    if not content:
        clog.error("Rebuild failed — check chunk data")
        raise HTTPException(status_code=400, detail="Rebuild failed — check chunk data")
    filename = req.chunks[0].get("filename", "file")
    clog.info(f"Rebuilt {len(content) / 1024:.1f} KB for {filename} — streaming back to caller")
    return StreamingResponse(
        io.BytesIO(content),
        media_type = "application/octet-stream",
        headers    = {"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ── S3 / Object Store ─────────────────────────────────────────────────────────

@app.get("/debug/s3", tags=["system"], summary="Diagnose S3 / Object Store connectivity", dependencies=[Depends(require_debug)])
def debug_s3():
    """
    Verifies S3/Object Store connectivity and shows the resolved credentials source.

    **Credential resolution order**
    1. `VCAP_SERVICES` (SAP BTP Object Store bound service — production)
    2. `s3_creds.json` file in the app directory (local dev)
    3. `AWS_ACCESS_KEY_ID` / `AWS_SECRET_ACCESS_KEY` / `S3_BUCKET` env vars (fallback)

    **Response fields**
    - `credentials.source` — which source was used (`VCAP_SERVICES`, `s3_creds.json`, or `env_vars`)
    - `credentials.bucket` / `region` / `endpoint_url` — resolved connection parameters
    - `ping.ok` — true if `head_bucket` succeeded
    - `ping.error` — S3 error code if the ping failed
    """
    result: dict = {}

    # ── Step 1: show raw VCAP_SERVICES structure (keys + instance names only, no secret values)
    vcap_raw = os.environ.get("VCAP_SERVICES")
    if vcap_raw:
        try:
            vcap = json.loads(vcap_raw)
            result["vcap_services"] = {
                "present": True,
                "top_level_keys": list(vcap.keys()),
                "services": {
                    label: [
                        {
                            "name": inst.get("name"),
                            "plan": inst.get("plan"),
                            "credential_keys": list(inst.get("credentials", {}).keys()),
                        }
                        for inst in instances
                    ]
                    for label, instances in vcap.items()
                },
            }
        except Exception as e:
            result["vcap_services"] = {"present": True, "parse_error": str(e)}
    else:
        result["vcap_services"] = {"present": False, "note": "VCAP_SERVICES env var not set — not running on CF or service not bound"}

    # ── Step 2: show which objectstore keys are being searched
    result["objectstore_keys_searched"] = list(s3._OBJECTSTORE_SERVICE_KEYS)

    # ── Step 3: resolve credentials and show source
    try:
        creds = s3._get_creds()
        result["credentials"] = {
            "source": (
                "VCAP_SERVICES" if vcap_raw and creds.get("access_key_id")
                else "s3_creds.json" if creds.get("access_key_id")
                else "env_vars (may be empty)"
            ),
            "access_key_id":     (creds.get("access_key_id") or "")[:6] + "***" if creds.get("access_key_id") else "(not set)",
            "secret_access_key": "***" if creds.get("secret_access_key") else "(not set)",
            "bucket":            creds.get("bucket", "(not set)"),
            "region":            creds.get("region", "(not set)"),
            "endpoint_url":      creds.get("endpoint_url", "(not set)"),
        }
    except Exception as e:
        result["credentials"] = {"error": type(e).__name__, "detail": str(e)}

    # ── Step 4: attempt connectivity ping
    try:
        ping = s3.ping()
        result["ping"] = ping
    except Exception as e:
        result["ping"] = {"ok": False, "error": type(e).__name__, "detail": str(e)}

    return result


@app.get("/debug/tmp", tags=["system"], summary="Browse the /tmp filesystem one level at a time", dependencies=[Depends(require_debug)])
def debug_tmp(path: str = Query("", description="Sub-path inside DOWNLOAD_PATH to browse, e.g. '20260829' or '20260829/runs'. Leave empty for the root.")):
    """
    Directory browser for the `DOWNLOAD_PATH` filesystem (typically `/tmp` on CF).

    Returns one level of directory contents at a time. Navigate into sub-directories using the
    `navigate` URL in each `dir` entry. Text-based files include a `preview` URL pointing to
    `GET /debug/tmp/preview`.

    **Response fields**
    - `current` — the path being listed
    - `parent` — query string to navigate up one level
    - `dir_count` / `file_count` / `total_size` — summary
    - `entries` — array of `{type, name, path, size, modified, navigate?, preview?}`

    Path is validated against `DOWNLOAD_PATH` to prevent directory traversal.

    **Previewable extensions:** txt, log, html, json, csv, yaml, yml, md, xml, ini, toml, env, cfg, sql, py, js, ts, java, cs, abap
    """
    base    = DOWNLOAD_PATH  # /tmp on CF
    current = os.path.normpath(os.path.join(base, path.lstrip("/\\")) if path else base)

    # Guard against path traversal outside DOWNLOAD_PATH
    base_norm = os.path.normpath(base)
    if os.path.commonpath([current, base_norm]) != base_norm:
        raise HTTPException(status_code=400, detail="Path outside allowed directory")
    if not os.path.exists(current):
        return {"path": path or "/", "exists": False, "entries": []}
    if not os.path.isdir(current):
        raise HTTPException(status_code=400, detail="Path is a file, not a directory")

    # Build parent path for navigation
    rel_current = os.path.relpath(current, base).replace("\\", "/")
    if rel_current == ".":
        rel_current = ""
    parent = os.path.relpath(os.path.dirname(current), base).replace("\\", "/")
    if parent == ".":
        parent = ""

    entries = []
    for name in sorted(os.listdir(current)):
        full = os.path.join(current, name)
        rel  = (f"{rel_current}/{name}" if rel_current else name).replace("\\", "/")
        if os.path.isdir(full):
            entries.append({
                "type":     "dir",
                "name":     name,
                "path":     rel,
                "navigate": f"/debug/tmp?path={rel}",
            })
        else:
            try:
                size  = os.path.getsize(full)
                mtime = datetime.utcfromtimestamp(os.path.getmtime(full)).strftime("%Y-%m-%d %H:%M:%S UTC")
            except OSError:
                size, mtime = None, "—"
            entries.append({
                "type":       "file",
                "name":       name,
                "path":       rel,
                "size_bytes": size,
                "size":       _fmt_size(size) if size is not None else "—",
                "modified":   mtime,
            })

    _PREVIEWABLE = {"txt", "log", "html", "htm", "json", "csv", "yaml", "yml",
                    "md", "xml", "ini", "toml", "env", "cfg", "sql", "py",
                    "js", "ts", "java", "cs", "abap"}

    dirs        = [e for e in entries if e["type"] == "dir"]
    files       = [e for e in entries if e["type"] == "file"]
    for e in files:
        ext = e["name"].rsplit(".", 1)[-1].lower() if "." in e["name"] else ""
        if ext in _PREVIEWABLE:
            e["preview"] = f"/debug/tmp/preview?path={e['path']}"
    total_bytes = sum(e["size_bytes"] for e in files if e["size_bytes"])
    return {
        "current":    f"/{rel_current}" if rel_current else "/",
        "parent":     f"?path={parent}" if parent else None,
        "dir_count":  len(dirs),
        "file_count": len(files),
        "total_size": _fmt_size(total_bytes),
        "entries":    entries,
    }


@app.get("/debug/tmp/preview", tags=["system"], summary="Preview a text file from /tmp", dependencies=[Depends(require_debug)])
def debug_tmp_preview(
    path:  str = Query(..., description="File path relative to DOWNLOAD_PATH, e.g. '20260829/sp_chunks_20260829_143000.jsonl'."),
    lines: int = Query(0, ge=0, description="Maximum number of lines to return. 0 (default) returns the full file."),
    tail:  bool = Query(False, description="When true, returns the last N lines instead of the first N. Requires lines > 0."),
):
    """
    Returns the text contents of a file inside `DOWNLOAD_PATH`.

    - **HTML / HTM files** — served as rendered HTML (browser-friendly)
    - **All other previewable types** — returned as a JSON response with `content`, `total_lines`, `shown_lines`, `truncated`

    Use `?lines=100` to preview the first 100 lines, or `?lines=50&tail=true` for the last 50.
    Full file is returned when `lines=0` (default).

    Path is validated against `DOWNLOAD_PATH` to prevent directory traversal.
    """
    base    = DOWNLOAD_PATH
    target  = os.path.normpath(os.path.join(base, path.lstrip("/\\")))

    base_norm = os.path.normpath(base)
    if os.path.commonpath([target, base_norm]) != base_norm:
        raise HTTPException(status_code=400, detail="Path outside allowed directory")
    if not os.path.exists(target):
        raise HTTPException(status_code=404, detail=f"File not found: {path}")
    if not os.path.isfile(target):
        raise HTTPException(status_code=400, detail="Path is a directory")

    ext = target.rsplit(".", 1)[-1].lower() if "." in target else ""

    # Serve HTML files as rendered HTML
    if ext in ("html", "htm"):
        with open(target, encoding="utf-8", errors="replace") as fh:
            content = fh.read()
        return HTMLResponse(content=content)

    # All other previewable types as plain text
    try:
        with open(target, encoding="utf-8", errors="replace") as fh:
            all_lines = fh.readlines()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Could not read file: {e}")

    if lines and lines > 0:
        selected = all_lines[-lines:] if tail else all_lines[:lines]
        truncated = len(all_lines) > lines
    else:
        selected  = all_lines
        truncated = False

    size = os.path.getsize(target)
    return {
        "path":        path,
        "size":        _fmt_size(size),
        "total_lines": len(all_lines),
        "shown_lines": len(selected),
        "truncated":   truncated,
        "content":     "".join(selected),
    }


@app.get("/s3/objects", tags=["s3"], summary="List objects in the S3 bucket")
def s3_list_objects(
    prefix:   str = Query("",   description="Key prefix filter to narrow results (e.g. 'runs/20260829'). Leave empty to list all objects."),
    max_keys: int = Query(1000, ge=1, le=10000, description="Maximum number of objects to return (1–10 000)."),
):
    """
    Lists objects in the configured SAP BTP Object Store / S3 bucket.

    Uses paginated `list_objects_v2` internally so all matching keys are returned regardless of
    bucket size (up to `max_keys`).

    **Response fields**
    - `bucket` — configured bucket name
    - `prefix` — the prefix filter applied
    - `count` — number of objects returned
    - `objects` — array of `{key, size, last_modified, etag, storage_class}`
    """
    try:
        objects = s3.list_objects(prefix=prefix, max_keys=max_keys)
        return {"bucket": s3.get_bucket_name(), "prefix": prefix, "count": len(objects), "objects": objects}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"{type(e).__name__}: {e}")


@app.get("/s3/object/{s3_key:path}/metadata", tags=["s3"], summary="Get metadata for an S3 object")
def s3_object_metadata(s3_key: str):
    """
    Returns head-object metadata for an S3 object without downloading it.

    **Response fields:** `content_type`, `content_length` (bytes), `last_modified` (ISO 8601), `etag`, `metadata` (user-defined key-value pairs).

    Returns **404** if the object does not exist.
    """
    if not s3.object_exists(s3_key):
        raise HTTPException(status_code=404, detail=f"Object not found: {s3_key}")
    try:
        return s3.get_metadata(s3_key)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"{type(e).__name__}: {e}")


@app.get("/s3/object/{s3_key:path}", tags=["s3"], summary="Stream-download an S3 object through the API")
def s3_download_object(s3_key: str):
    """
    Downloads an S3 object and streams it back to the caller as an `application/octet-stream` attachment.

    The filename in the `Content-Disposition` header is the last path segment of `s3_key`.
    Returns **404** if the object does not exist.

    For large files, consider `GET /s3/presigned/{s3_key}` to get a direct presigned URL instead
    — it avoids proxying the bytes through the API.
    """
    if not s3.object_exists(s3_key):
        raise HTTPException(status_code=404, detail=f"Object not found: {s3_key}")
    try:
        filename = s3_key.rsplit("/", 1)[-1]
        return StreamingResponse(
            s3.stream_object(s3_key),
            media_type = "application/octet-stream",
            headers    = {"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"{type(e).__name__}: {e}")


@app.get("/s3/presigned/{s3_key:path}", tags=["s3"], summary="Generate a presigned GET URL for direct S3 download")
def s3_presigned_download(
    s3_key: str,
    expiry: int = Query(3600, ge=60, le=86400, description="URL lifetime in seconds. Min 60 s, max 86 400 s (24 h)."),
):
    """
    Generates a time-limited presigned GET URL that allows anyone with the URL to download
    the S3 object directly — without going through this API and without AWS credentials.

    **Response fields**
    - `s3_key` — the object key
    - `url` — the presigned URL (valid for `expires_in` seconds)
    - `expires_in` — lifetime in seconds

    Use this for large files or for sharing download links with clients that should not access
    the API directly.
    """
    try:
        url = s3.get_presigned_url(s3_key, expiry=expiry)
        return {"s3_key": s3_key, "url": url, "expires_in": expiry}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"{type(e).__name__}: {e}")


@app.post("/s3/presigned-upload", tags=["s3"], summary="Generate a presigned POST policy for direct browser-to-S3 upload")
def s3_presigned_upload_url(req: S3PresignedUploadRequest):
    """
    Generates a presigned POST policy that allows a browser or HTTP client to upload a file
    **directly to S3** without routing bytes through this API.

    **Response shape:** `{"url": "https://...", "fields": {"key": "...", "Content-Type": "...", ...}}`

    The client must submit a multipart POST to `url` with all `fields` included as form fields,
    followed by the file under the `file` key. The upload is rejected by S3 if:
    - `Content-Type` does not match the value in `fields`
    - File size exceeds `max_size_mb`
    - The URL has expired

    This approach is preferred over `POST /s3/upload-storage-file` when the source bytes are
    already on the client (avoids double-transfer through the API).
    """
    try:
        return s3.get_presigned_upload_url(
            req.s3_key,
            expiry         = req.expiry,
            content_type   = req.content_type,
            max_size_bytes = req.max_size_mb * 1024 * 1024,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"{type(e).__name__}: {e}")


@app.delete("/s3/object/{s3_key:path}", tags=["s3"], summary="Delete a single S3 object", dependencies=[Depends(require_destructive)])
def s3_delete_object(s3_key: str):
    """
    Permanently deletes a single object from the S3 bucket.

    Returns `{"deleted": "<s3_key>"}` on success.
    No error is raised if the object does not exist (S3 DELETE is idempotent).
    """
    try:
        s3.delete_object(s3_key)
        return {"deleted": s3_key}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"{type(e).__name__}: {e}")


@app.delete("/s3/objects", tags=["s3"], summary="Bulk delete up to 1 000 S3 objects", dependencies=[Depends(require_destructive)])
def s3_bulk_delete(req: S3BulkDeleteRequest):
    """
    Deletes up to 1 000 S3 objects in a single S3 API call (`delete_objects`).

    **Response fields** (raw S3 response):
    - `Deleted` — list of successfully deleted keys
    - `Errors` — list of keys that could not be deleted with error codes

    Use `GET /s3/objects?prefix=<prefix>` to list keys first, then pass them here.
    """
    try:
        return s3.delete_objects(req.keys)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"{type(e).__name__}: {e}")


@app.post("/s3/upload-storage-file", tags=["s3"], summary="Upload a local output file to S3")
def s3_upload_storage_file(req: S3UploadStorageRequest):
    """
    Uploads a single file from `DOWNLOAD_PATH/<day>/<filename>` to the S3 bucket.

    Use the `date` and `filename` fields from `GET /storage` to build the request.
    The `s3_key` defaults to `<day>/<filename>` if not specified.

    **Response fields:** `uploaded` (s3_key), `source` (local relative path), `size_bytes`

    **Typical use case:** persist a specific vectorize output file (e.g. `sp_vectors_*.jsonl` or
    `sp_analysis_*.html`) to S3 before the CF `/tmp` volume is wiped on dyno restart.

    For bulk upload of an entire date folder use `POST /s3/upload-storage-folder/{day}`.
    """
    resolved = _safe_resolve(req.day, req.filename)
    if not os.path.isfile(resolved):
        raise HTTPException(status_code=404, detail=f"{req.day}/{req.filename} not found in local storage")
    s3_key = req.s3_key or f"{req.day}/{req.filename}"
    try:
        s3.upload_file(resolved, s3_key, content_type=req.content_type)
        log.info(f"Uploaded {req.day}/{req.filename} → s3:{s3_key}")
        return {"uploaded": s3_key, "source": f"{req.day}/{req.filename}", "size_bytes": os.path.getsize(resolved)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"{type(e).__name__}: {e}")


@app.post("/s3/upload-storage-folder/{day}", tags=["s3"], summary="Bulk upload an entire date folder to S3")
def s3_upload_storage_folder(
    day:       str,
    s3_prefix: str = Query("", description="S3 key prefix for all uploaded files (e.g. 'backups/20260829'). Defaults to the day string itself."),
):
    """
    Uploads every file inside `DOWNLOAD_PATH/<day>/` to the S3 bucket.

    Each file is uploaded to `<s3_prefix>/<filename>` (prefix defaults to `<day>`).

    **Response fields**
    - `day` — the source folder
    - `s3_prefix` — the prefix used for all S3 keys
    - `uploaded` — count of successfully uploaded files
    - `errors` — count of failed uploads
    - `files` — array of `{file, s3_key, size_bytes}` for successes
    - `failed` — array of `{file, error}` for failures

    **Typical use case:** back up an entire vectorize run (all chunks, vectors, reports, logs)
    to S3 in one call before restarting the CF app or freeing disk space.
    """
    safe_root = os.path.realpath(DOWNLOAD_PATH)
    day_path  = os.path.realpath(os.path.join(DOWNLOAD_PATH, day))
    if not day_path.startswith(safe_root + os.sep):
        raise HTTPException(status_code=400, detail="Invalid path")
    if not os.path.isdir(day_path):
        raise HTTPException(status_code=404, detail=f"Folder {day} not found")

    prefix   = s3_prefix or day
    uploaded = []
    errors   = []
    for fname in sorted(os.listdir(day_path)):
        fpath = os.path.join(day_path, fname)
        if not os.path.isfile(fpath):
            continue
        s3_key = f"{prefix}/{fname}"
        try:
            s3.upload_file(fpath, s3_key)
            uploaded.append({"file": fname, "s3_key": s3_key, "size_bytes": os.path.getsize(fpath)})
        except Exception as e:
            errors.append({"file": fname, "error": str(e)})

    log.info(f"Bulk S3 upload {day}: {len(uploaded)} ok, {len(errors)} failed")
    return {
        "day":       day,
        "s3_prefix": prefix,
        "uploaded":  len(uploaded),
        "errors":    len(errors),
        "files":     uploaded,
        "failed":    errors,
    }


# ── /vectorize/tmp endpoints ──────────────────────────────────────────────────

_TMP_SUPPORTED = {
    "docx","doc","xlsx","xlsm","xls","pdf","pptx","pptm",
    "txt","csv","json","md","rtf","msg","html","htm","py","js","ts","yaml","yml","sh",
}


@app.post(
    "/vectorize/tmp",
    response_model = VectorizeResponse,
    tags           = ["rag"],
    summary        = "Vectorize all files under a /tmp folder path",
)
def vectorize_tmp(req: TmpVectorizeRequest):
    """
    Reads every supported file under the given **local folder**, extracts text,
    chunks, embeds with `all-MiniLM-L6-v2`, and upserts chunks + vectors into HANA.

    The folder must be an absolute path accessible to the running CF instance (e.g. `/tmp/uploads`).

    **Change detection** — unless `force=true`, files whose `last_modified` timestamp
    already matches the value stored in HANA are skipped.

    **Supported file types** — docx, xlsx, pdf, pptx, txt, csv, json, md, rtf, msg, html, py, js, ts, yaml, sh

    **Body example**
    ```json
    {
      "folder": "/tmp/uploads",
      "force": false,
      "extensions": ["pdf", "docx"]
    }
    ```
    `extensions` filters to only those types; omit or pass `[]` for all supported types.
    """
    clog = _call_logger("vectorize_tmp")
    clog.info(f"POST /vectorize/tmp — folder={req.folder!r} force={req.force} extensions={req.extensions}")

    ext_filter = {e.lower().lstrip(".") for e in req.extensions} if req.extensions else None

    if not os.path.isdir(req.folder):
        raise HTTPException(status_code=404, detail=f"Folder not found or not a directory: {req.folder}")

    all_paths = []
    for fname in sorted(os.listdir(req.folder)):
        fpath = os.path.join(req.folder, fname)
        if not os.path.isfile(fpath):
            continue
        ext = fname.rsplit(".", 1)[-1].lower() if "." in fname else ""
        if ext not in _TMP_SUPPORTED:
            continue
        if ext_filter and ext not in ext_filter:
            continue
        all_paths.append(fpath)

    clog.info(f"Found {len(all_paths)} supported files in {req.folder!r}")

    if not all_paths:
        raise HTTPException(
            status_code=404,
            detail=f"No supported files found in '{req.folder}'. Supported types: {sorted(_TMP_SUPPORTED)}",
        )

    if req.force:
        to_process        = all_paths
        skipped_unchanged = 0
        clog.info("force=true — processing all files regardless of HANA state")
    else:
        items_for_filter = [_tmp_file_to_item(p) for p in all_paths]
        filtered, skipped_unchanged = _filter_stale_items(items_for_filter, logger=clog)
        processed_ids = {item["id"] for item in filtered}
        to_process    = [p for p in all_paths if hashlib.md5(p.encode()).hexdigest() in processed_ids]

    clog.info(f"{len(to_process)} to vectorize, {skipped_unchanged} unchanged (skipped)")

    if not to_process:
        clog.info("All /tmp files already up-to-date in HANA")
        return VectorizeResponse(
            count_chunks      = 0,
            count_files       = 0,
            skipped_unchanged = skipped_unchanged,
            saved_to          = [],
            db_ingest         = {"message": "all files up-to-date"},
        )

    try:
        n_chunks, n_files, path_chunks, path_raw, path_vectors, path_root, path_html, path_log, ingest_results = \
            _download_and_vectorize_tmp(to_process, logger=clog)
    except Exception as e:
        clog.error(f"TMP vectorization failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"TMP vectorization failed: {e}")

    clog.info(f"Complete — {n_files} files, {n_chunks} chunks, {skipped_unchanged} unchanged")
    return VectorizeResponse(
        count_chunks      = n_chunks,
        count_files       = n_files,
        skipped_unchanged = skipped_unchanged,
        saved_to          = [p for p in [path_chunks, path_raw, path_vectors, path_root, path_html, path_log] if p],
        db_ingest         = {"per_file_ingest": ingest_results},
    )


@app.get(
    "/vectorize/tmp/{folder_path:path}",
    response_model = VectorizeResponse,
    tags           = ["rag"],
    summary        = "Vectorize all files under a /tmp folder (GET, path-based)",
)
def vectorize_tmp_folder(
    folder_path: str,
    force:      bool = Query(False, description="Re-vectorize all files even if unchanged in HANA"),
    extensions: str  = Query("",    description="Comma-separated extension whitelist, e.g. 'pdf,docx'. Empty = all supported types."),
):
    """
    Same pipeline as `POST /vectorize/tmp` but accepts the folder as a URL path segment.

    **Example:** `GET /vectorize/tmp/uploads/knowledge-base`

    The path is appended to `/tmp/` automatically — do **not** include `/tmp/` in the URL.

    Query params:
    - `force` — re-vectorize all files regardless of HANA state
    - `extensions` — comma-separated whitelist, e.g. `pdf,docx`
    """
    abs_folder = os.path.join("/tmp", folder_path.lstrip("/"))
    ext_list   = [e.strip() for e in extensions.split(",") if e.strip()] if extensions else []
    req        = TmpVectorizeRequest(folder=abs_folder, force=force, extensions=ext_list)
    return vectorize_tmp(req)


@app.post(
    "/vectorize/tmp/item",
    response_model = S3VectorizeItemResponse,
    tags           = ["rag"],
    summary        = "Vectorize a single file from the local /tmp filesystem",
)
def vectorize_tmp_item(req: TmpVectorizeItemRequest):
    """
    Fetches and vectorizes a single file from the local filesystem by its absolute path.

    **Pipeline steps**
    1. Verify the file exists and get `last_modified` from `os.stat()`
    2. Change detection against HANA — return immediately with `skipped_unchanged=1` if up-to-date
    3. Read bytes → extract → chunk → embed → upsert to HANA
    4. Write JSONL output files, generate HTML report, upload artifacts to S3

    **Body example**
    ```json
    { "path": "/tmp/uploads/report.pdf", "force": false }
    ```
    """
    clog     = _call_logger("vectorize_tmp_item")
    abs_path = req.path
    filename = os.path.basename(abs_path)
    ext      = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    clog.info(f"POST /vectorize/tmp/item — path={abs_path!r} force={req.force}")

    if ext not in _TMP_SUPPORTED:
        raise HTTPException(
            status_code=422,
            detail=f"File extension '.{ext}' is not supported. Supported: {sorted(_TMP_SUPPORTED)}",
        )

    if not os.path.isfile(abs_path):
        raise HTTPException(status_code=404, detail=f"File not found: {abs_path}")

    item = _tmp_file_to_item(abs_path)

    if not req.force:
        filtered, skipped = _filter_stale_items([item], logger=clog)
        if skipped:
            clog.info(f"Skipping {abs_path!r} — already up-to-date in HANA")
            return S3VectorizeItemResponse(
                key               = abs_path,
                filename          = filename,
                count_chunks      = 0,
                skipped_unchanged = 1,
                saved_to          = [],
                db_ingest         = {"message": "already up-to-date"},
            )

    try:
        n_chunks, n_files, path_chunks, path_raw, path_vectors, path_root, path_html, path_log, ingest_results = \
            _download_and_vectorize_tmp([abs_path], logger=clog)
    except Exception as e:
        clog.error(f"TMP item vectorization failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"TMP item vectorization failed: {e}")

    clog.info(f"Complete — path={abs_path!r} chunks={n_chunks}")
    return S3VectorizeItemResponse(
        key               = abs_path,
        filename          = filename,
        count_chunks      = n_chunks,
        skipped_unchanged = 0,
        saved_to          = [p for p in [path_chunks, path_raw, path_vectors, path_root, path_html, path_log] if p],
        db_ingest         = {"per_file_ingest": ingest_results},
    )


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    uvicorn.run("api:app", host="0.0.0.0", port=int(os.getenv("PORT", 8000)), reload=False, timeout_keep_alive=600)
