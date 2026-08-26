import os
import json
import logging
import threading
from datetime import datetime

from hdbcli import dbapi
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, Field

log = logging.getLogger("hana_db")

_conn      = None
_conn_lock = threading.Lock()
_schema    = ""


def _t(table: str) -> str:
    """Return schema-qualified table name: schema.TABLE or just TABLE if no schema."""
    return f"{_schema}.{table}" if _schema else table


# ── Credentials ───────────────────────────────────────────────────────────────

def _get_creds() -> dict:
    # 1 — CF bound service (production)
    vcap_raw = os.environ.get("VCAP_SERVICES")
    if vcap_raw:
        services = json.loads(vcap_raw)
        for key in ("hana", "hanatrial", "hana-cloud"):
            if key in services:
                return services[key][0]["credentials"]
        for instances in services.values():
            for inst in instances:
                creds = inst.get("credentials", {})
                if "host" in creds and "port" in creds:
                    return creds

    # 2 — hana_creds.json file (local dev)
    creds_file = os.path.join(os.path.dirname(__file__), "hana_creds.json")
    if os.path.isfile(creds_file):
        with open(creds_file, encoding="utf-8") as fh:
            data = json.load(fh)
        if data.get("host") and data["host"] != "your-hana-instance.hanacloud.ondemand.com":
            log.info("HANA credentials loaded from hana_creds.json")
            return {
                "host":                  data["host"],
                "port":                  int(data.get("port", 443)),
                "user":                  data["user"],
                "password":              data["password"],
                "schema":                data.get("schema", ""),
                "encrypt":               data.get("encrypt", True),
                "sslValidateCertificate": data.get("sslValidateCertificate", True),
            }

    # 3 — environment variables (local dev fallback)
    return {
        "host":                  os.getenv("HANA_HOST", ""),
        "port":                  int(os.getenv("HANA_PORT", 443)),
        "user":                  os.getenv("HANA_USER", ""),
        "password":              os.getenv("HANA_PASSWORD", ""),
        "schema":                os.getenv("HANA_SCHEMA", ""),
        "encrypt":               True,
        "sslValidateCertificate": True,
    }


# ── Connection ────────────────────────────────────────────────────────────────

def _connect() -> dbapi.Connection:
    global _conn, _schema
    with _conn_lock:
        try:
            if _conn is not None and _conn.isconnected():
                return _conn
        except Exception:
            _conn = None
        creds = _get_creds()
        if not creds.get("host"):
            raise RuntimeError("HANA credentials not found in VCAP_SERVICES or env vars")

        print(f"[DB] Connecting to HANA — host={creds.get('host')} port={creds.get('port')} user={creds.get('user')} schema={creds.get('schema', '(none)')}")

        connect_kwargs = {
            "address":               creds["host"],
            "port":                  int(creds["port"]),
            "user":                  creds["user"],
            "password":              creds["password"],
            "encrypt":               creds.get("encrypt", True),
            "sslValidateCertificate": creds.get("sslValidateCertificate", True),
        }
        if creds.get("schema"):
            connect_kwargs["currentSchema"] = creds["schema"]
            _schema = creds["schema"]

        try:
            _conn = dbapi.connect(**connect_kwargs)
            print(f"[DB] HANA connection established — schema={_schema}")
            log.info(f"HANA connected — {creds['host']}:{creds['port']} schema={_schema}")
        except Exception as e:
            print(f"[DB] HANA connection FAILED — {e}")
            log.error(f"HANA connection failed: {e}")
            raise

        return _conn


# ── Schema ────────────────────────────────────────────────────────────────────

def create_tables() -> dict:
    conn   = _connect()
    cur    = conn.cursor()
    status = {}
    try:
        cur.execute("SELECT TABLE_NAME FROM SYS.TABLES WHERE SCHEMA_NAME = CURRENT_SCHEMA")
        existing = {row[0] for row in cur.fetchall()}

        # Keys are plain (unqualified) names — used to check SYS.TABLES.
        # DDL uses _t() so the CREATE TABLE targets the correct schema.
        tables = {
            "CATALYSTPLATFORM_BRAIN_FILES": f"""
                CREATE TABLE {_t("CATALYSTPLATFORM_BRAIN_FILES")} (
                    item_id       NVARCHAR(200) PRIMARY KEY,
                    filename      NVARCHAR(500),
                    file_type     NVARCHAR(200),
                    extension     NVARCHAR(20),
                    size_text     NVARCHAR(50),
                    size_bytes    BIGINT,
                    web_url       NVARCHAR(1000),
                    doc_path      NVARCHAR(1000),
                    last_modified NVARCHAR(50),
                    created_date  NVARCHAR(50),
                    created_by    NVARCHAR(500),
                    modified_by   NVARCHAR(500),
                    shared        NVARCHAR(100),
                    version_count INT,
                    indexed_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    insert_dt     TIMESTAMP
                )
            """,
            "CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS": f"""
                CREATE TABLE {_t("CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS")} (
                    item_id    NVARCHAR(200),
                    chunk_id   INT,
                    chunk_data NCLOB,
                    filename   NVARCHAR(500),
                    insert_dt  TIMESTAMP,
                    PRIMARY KEY (item_id, chunk_id)
                )
            """,
            "CATALYSTPLATFORM_BRAIN_VECTORS": f"""
                CREATE TABLE {_t("CATALYSTPLATFORM_BRAIN_VECTORS")} (
                    item_id     NVARCHAR(200),
                    chunk_id    INT,
                    vector_data REAL_VECTOR(384),
                    filename    NVARCHAR(500),
                    insert_dt   TIMESTAMP,
                    PRIMARY KEY (item_id, chunk_id)
                )
            """,
            "CATALYSTPLATFORM_BRAIN_FILES_VERSIONS": f"""
                CREATE TABLE {_t("CATALYSTPLATFORM_BRAIN_FILES_VERSIONS")} (
                    item_id          NVARCHAR(200),
                    version_id       NVARCHAR(50),
                    file_size_text   NVARCHAR(50),
                    size_bytes       BIGINT,
                    last_modified_dt NVARCHAR(50),
                    lastmodified_by  NVARCHAR(500),
                    download_url     NCLOB,
                    version_count    INT,
                    filename         NVARCHAR(500),
                    insert_dt        TIMESTAMP,
                    PRIMARY KEY (item_id, version_id)
                )
            """,
            "CATALYSTPLATFORM_BRAIN_FILES_CHUNK": f"""
                CREATE TABLE {_t("CATALYSTPLATFORM_BRAIN_FILES_CHUNK")} (
                    item_id     NVARCHAR(200) PRIMARY KEY,
                    filename    NVARCHAR(500),
                    chunk_count INT,
                    insert_dt   TIMESTAMP
                )
            """,
        }

        for table_name, ddl in tables.items():
            if table_name in existing:
                status[table_name] = "already exists"
                log.info(f"{_t(table_name)} — already exists")
                print(f"[DB] TABLE ALREADY EXISTS: {_t(table_name)}")
            else:
                cur.execute(ddl)
                status[table_name] = "created"
                log.info(f"{_t(table_name)} — created")
                print(f"[DB] TABLE CREATED: {_t(table_name)}")

        conn.commit()
        log.info(f"HANA tables ready — {status}")
        print(f"[DB] All tables ready: {status}")
        return status
    finally:
        cur.close()


# ── Ingest ────────────────────────────────────────────────────────────────────

def ingest(chunks_data: list, vectors_data: list, items: list) -> dict:
    conn      = _connect()
    cur       = conn.cursor()
    files_map = {item["id"]: item for item in items if item.get("id")}
    now       = datetime.utcnow()
    try:
        for item_id, item in files_map.items():
            cur.execute(
                f"UPSERT {_t('CATALYSTPLATFORM_BRAIN_FILES')} "
                "(item_id, filename, file_type, extension, size_text, size_bytes, web_url, doc_path, "
                " last_modified, created_date, created_by, modified_by, shared, version_count, insert_dt) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) WITH PRIMARY KEY",
                (
                    item_id,
                    item.get("name", ""),
                    item.get("type", ""),
                    item.get("extension", ""),
                    item.get("size", ""),
                    item.get("size_bytes"),
                    item.get("web_url", ""),
                    item.get("doc_path", ""),
                    item.get("last_modified_date", ""),
                    item.get("created_date", ""),
                    item.get("created_by", ""),
                    item.get("modified_by", ""),
                    item.get("shared", ""),
                    item.get("version_count"),
                    now,
                )
            )

        for c in chunks_data:
            cur.execute(
                f"UPSERT {_t('CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS')} "
                "(item_id, chunk_id, chunk_data, filename, insert_dt) "
                "VALUES (?, ?, ?, ?, ?) WITH PRIMARY KEY",
                (c["item_id"], c["chunk_id"], c["chunk_data"], c.get("filename", ""), now)
            )

        for v in vectors_data:
            vec_str = "[" + ",".join(f"{x:.8f}" for x in v["vector_data"]) + "]"
            cur.execute(
                f"UPSERT {_t('CATALYSTPLATFORM_BRAIN_VECTORS')} "
                "(item_id, chunk_id, vector_data, filename, insert_dt) "
                "VALUES (?, ?, TO_REAL_VECTOR(?), ?, ?) WITH PRIMARY KEY",
                (v["item_id"], v["chunk_id"], vec_str, v.get("filename", ""), now)
            )

        conn.commit()
        result = {"files": len(files_map), "chunks": len(chunks_data), "vectors": len(vectors_data)}
        log.info(f"Ingested — {result}")
        return result
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()


def ingest_versions(versions_data: list) -> dict:
    if not versions_data:
        return {"versions": 0}
    conn = _connect()
    cur  = conn.cursor()
    now  = datetime.utcnow()
    try:
        for v in versions_data:
            cur.execute(
                f"UPSERT {_t('CATALYSTPLATFORM_BRAIN_FILES_VERSIONS')} "
                "(item_id, version_id, file_size_text, size_bytes, last_modified_dt, "
                " lastmodified_by, download_url, version_count, filename, insert_dt) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) WITH PRIMARY KEY",
                (
                    v.get("item_id", ""),
                    str(v.get("version_id", "")),
                    v.get("fileSize", ""),
                    v.get("size"),
                    v.get("lastModifiedDateTime", ""),
                    v.get("lastmodified_by"),
                    v.get("download_url", ""),
                    v.get("version_count"),
                    v.get("filename", ""),
                    now,
                )
            )
        conn.commit()
        log.info(f"Ingested {len(versions_data)} version records")
        return {"versions": len(versions_data)}
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()


def ingest_chunk_summary(raw_chunks_data: list) -> dict:
    if not raw_chunks_data:
        return {"chunk_summaries": 0}
    conn = _connect()
    cur  = conn.cursor()
    now  = datetime.utcnow()
    try:
        for r in raw_chunks_data:
            cur.execute(
                f"UPSERT {_t('CATALYSTPLATFORM_BRAIN_FILES_CHUNK')} "
                "(item_id, filename, chunk_count, insert_dt) "
                "VALUES (?, ?, ?, ?) WITH PRIMARY KEY",
                (
                    r.get("item_id", ""),
                    r.get("filename", ""),
                    r.get("chunk_count"),
                    now,
                )
            )
        conn.commit()
        log.info(f"Ingested {len(raw_chunks_data)} chunk summaries")
        return {"chunk_summaries": len(raw_chunks_data)}
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()


# ── Query ─────────────────────────────────────────────────────────────────────

def query(vector: list, top_k: int = 5) -> list:
    vec_str = "[" + ",".join(f"{x:.8f}" for x in vector) + "]"
    conn    = _connect()
    cur     = conn.cursor()
    try:
        cur.execute(f"""
            SELECT TOP {top_k}
                v.item_id, f.filename, v.chunk_id, c.chunk_data,
                COSINE_SIMILARITY(v.vector_data, TO_REAL_VECTOR(?)) AS score
            FROM {_t("CATALYSTPLATFORM_BRAIN_VECTORS")} v
            JOIN {_t("CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS")} c ON v.item_id = c.item_id AND v.chunk_id = c.chunk_id
            JOIN {_t("CATALYSTPLATFORM_BRAIN_FILES")} f ON v.item_id = f.item_id
            ORDER BY score DESC
        """, (vec_str,))
        return [
            {
                "item_id":    row[0],
                "filename":   row[1],
                "chunk_id":   row[2],
                "chunk_data": row[3],
                "score":      float(row[4]),
            }
            for row in cur.fetchall()
        ]
    finally:
        cur.close()


# ── CRUD ──────────────────────────────────────────────────────────────────────

def get_files() -> list:
    conn = _connect()
    cur  = conn.cursor()
    try:
        cur.execute(
            f"SELECT item_id, filename, file_type, extension, size_text, size_bytes, "
            f"web_url, doc_path, last_modified, created_date, created_by, modified_by, "
            f"shared, version_count, indexed_at, insert_dt "
            f"FROM {_t('CATALYSTPLATFORM_BRAIN_FILES')} ORDER BY indexed_at DESC"
        )
        cols = [d[0].lower() for d in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()]
    finally:
        cur.close()


def get_chunks(item_id: str) -> list:
    conn = _connect()
    cur  = conn.cursor()
    try:
        cur.execute(
            f"SELECT chunk_id, chunk_data, filename, insert_dt "
            f"FROM {_t('CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS')} WHERE item_id = ? ORDER BY chunk_id",
            (item_id,)
        )
        cols = [d[0].lower() for d in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()]
    finally:
        cur.close()


def get_versions(item_id: str = None) -> list:
    conn = _connect()
    cur  = conn.cursor()
    try:
        if item_id:
            cur.execute(
                f"SELECT item_id, version_id, file_size_text, size_bytes, last_modified_dt, "
                f"lastmodified_by, version_count, filename, insert_dt "
                f"FROM {_t('CATALYSTPLATFORM_BRAIN_FILES_VERSIONS')} WHERE item_id = ? ORDER BY version_id",
                (item_id,)
            )
        else:
            cur.execute(
                f"SELECT item_id, version_id, file_size_text, size_bytes, last_modified_dt, "
                f"lastmodified_by, version_count, filename, insert_dt "
                f"FROM {_t('CATALYSTPLATFORM_BRAIN_FILES_VERSIONS')} ORDER BY item_id, version_id"
            )
        cols = [d[0].lower() for d in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()]
    finally:
        cur.close()


def get_chunk_summary(item_id: str = None) -> list:
    conn = _connect()
    cur  = conn.cursor()
    try:
        if item_id:
            cur.execute(
                f"SELECT item_id, filename, chunk_count, insert_dt "
                f"FROM {_t('CATALYSTPLATFORM_BRAIN_FILES_CHUNK')} WHERE item_id = ?",
                (item_id,)
            )
        else:
            cur.execute(
                f"SELECT item_id, filename, chunk_count, insert_dt "
                f"FROM {_t('CATALYSTPLATFORM_BRAIN_FILES_CHUNK')} ORDER BY filename"
            )
        cols = [d[0].lower() for d in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()]
    finally:
        cur.close()


def delete_item(item_id: str) -> dict:
    conn = _connect()
    cur  = conn.cursor()
    try:
        cur.execute(f"DELETE FROM {_t('CATALYSTPLATFORM_BRAIN_FILES_VERSIONS')} WHERE item_id = ?", (item_id,))
        cur.execute(f"DELETE FROM {_t('CATALYSTPLATFORM_BRAIN_FILES_CHUNK')}    WHERE item_id = ?", (item_id,))
        cur.execute(f"DELETE FROM {_t('CATALYSTPLATFORM_BRAIN_VECTORS')}        WHERE item_id = ?", (item_id,))
        cur.execute(f"DELETE FROM {_t('CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS')}    WHERE item_id = ?", (item_id,))
        cur.execute(f"DELETE FROM {_t('CATALYSTPLATFORM_BRAIN_FILES')}          WHERE item_id = ?", (item_id,))
        conn.commit()
        log.info(f"Deleted item {item_id}")
        return {"deleted": item_id}
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()


# ── Router ────────────────────────────────────────────────────────────────────

router = APIRouter(prefix="/db", tags=["database"])


class QueryRequest(BaseModel):
    text:  str = Field(...,  description="Natural language query to search for")
    top_k: int = Field(5, ge=1, le=50, description="Number of results to return")


@router.post("/query")
def db_query(req: QueryRequest):
    from api import _get_embed_model
    try:
        vector  = _get_embed_model().encode(req.text).tolist()
        results = query(vector, top_k=req.top_k)
        return {"query": req.text, "count": len(results), "results": results}
    except Exception as e:
        log.error(f"Query failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/files")
def db_files():
    try:
        files = get_files()
        return {"count": len(files), "files": files}
    except Exception as e:
        log.error(f"DB files failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/chunks/{item_id}")
def db_chunks(item_id: str):
    try:
        chunks = get_chunks(item_id)
        if not chunks:
            raise HTTPException(status_code=404, detail=f"No chunks found for item_id: {item_id}")
        return {"item_id": item_id, "count": len(chunks), "chunks": chunks}
    except HTTPException:
        raise
    except Exception as e:
        log.error(f"DB chunks failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/versions")
def db_versions_all():
    try:
        rows = get_versions()
        return {"count": len(rows), "versions": rows}
    except Exception as e:
        log.error(f"DB versions failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/versions/{item_id}")
def db_versions(item_id: str):
    try:
        rows = get_versions(item_id)
        return {"item_id": item_id, "count": len(rows), "versions": rows}
    except Exception as e:
        log.error(f"DB versions failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/chunk-summary")
def db_chunk_summary_all():
    try:
        rows = get_chunk_summary()
        return {"count": len(rows), "chunk_summaries": rows}
    except Exception as e:
        log.error(f"DB chunk summary failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/chunk-summary/{item_id}")
def db_chunk_summary(item_id: str):
    try:
        rows = get_chunk_summary(item_id)
        if not rows:
            raise HTTPException(status_code=404, detail=f"No chunk summary for item_id: {item_id}")
        return {"item_id": item_id, "chunk_summary": rows[0]}
    except HTTPException:
        raise
    except Exception as e:
        log.error(f"DB chunk summary failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.delete("/item/{item_id}")
def db_delete(item_id: str):
    try:
        return delete_item(item_id)
    except Exception as e:
        log.error(f"DB delete failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))
