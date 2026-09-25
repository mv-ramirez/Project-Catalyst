import os
import json
import logging
import threading
from datetime import datetime

from hdbcli import dbapi
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel, Field

log = logging.getLogger("hana_db")

_conn      = None
_conn_lock = threading.Lock()
_schema    = ""
_col_cache: dict = {}   # table_name.upper() → frozenset of lowercase column names


def _t(table: str) -> str:
    """Return schema-qualified table name: schema.TABLE or just TABLE if no schema."""
    return f"{_schema}.{table}" if _schema else table


# ── Table names ───────────────────────────────────────────────────────────────
# Two physical table sets exist during the HDI migration:
#
#   "legacy" — CATALYSTPLATFORM_BRAIN_*    created and owned by this module.
#   "hdi"    — CATALYSTPLATFORM_TB_BRAIN_* owned by the Node HDI db module
#              (BTP-Catalyst-Platform/db/src/*.hdbmigrationtable).
#
# Which set is read from, and which are written to, is configured just below.
# Both default to "legacy", so a plain `cf push` changes nothing: Python keeps
# creating and using the tables it always has.
#
# Python issues DDL against the legacy set ONLY. create_tables(), migrate_tables()
# and drop_tables() refuse to run when the HDI set is active; those tables belong
# to HDI and Python creating objects inside a container HDI also owns is what
# caused the Sep-10 loss.
#
# Renaming an HDI table is a two-repo change: the .hdbmigrationtable file and the
# constant here move together, in the same commit. Always go through these
# constants; never inline a table name.

LEGACY_TABLES = {
    "sp_files":       "CATALYSTPLATFORM_BRAIN_SP_FILES",
    "s3_files":       "CATALYSTPLATFORM_BRAIN_S3_FILES",
    "text_chunks":    "CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS",
    "vectors":        "CATALYSTPLATFORM_BRAIN_VECTORS",
    "files_chunk":    "CATALYSTPLATFORM_BRAIN_FILES_CHUNK",
    "file_registry":  "CATALYSTPLATFORM_BRAIN_FILE_REGISTRY",
    "sp_versions":    "CATALYSTPLATFORM_BRAIN_SP_FILES_VERSIONS",
}

HDI_TABLES = {
    "sp_files":       "CATALYSTPLATFORM_TB_BRAIN_SP_FILES",
    "s3_files":       "CATALYSTPLATFORM_TB_BRAIN_S3_FILES",
    "text_chunks":    "CATALYSTPLATFORM_TB_BRAIN_TEXT_CHUNKS",
    "vectors":        "CATALYSTPLATFORM_TB_BRAIN_VECTORS",
    "files_chunk":    "CATALYSTPLATFORM_TB_BRAIN_FILES_CHUNK",
    "file_registry":  "CATALYSTPLATFORM_TB_BRAIN_FILE_REGISTRY",
    "sp_versions":    "CATALYSTPLATFORM_TB_BRAIN_SP_FILES_VERSIONS",
}

TABLE_SETS = {"legacy": LEGACY_TABLES, "hdi": HDI_TABLES}

# Reads and writes are configured separately, because during the migration they
# genuinely are two decisions:
#
#   BRAIN_WRITE_SET = legacy | hdi | both   (default: legacy)
#   BRAIN_READ_SET  = legacy | hdi          (default: legacy)
#
# Setting BRAIN_WRITE_SET=both makes every ingest land in BOTH table sets, so the
# HDI tables fill with real data from real API calls while the legacy tables stay
# authoritative. That is how the insert path gets proven against the new schema
# without a cutover: run the normal vectorize endpoints, then diff the two sets.
#
# Writes to the READ set are primary — a failure there raises, exactly as before.
# Writes to the other set are best-effort: a failure is logged and returned in the
# ingest result under "errors", but never breaks the live path. Tests assert that
# "errors" is empty; production keeps serving if the new tables have a problem.
BRAIN_READ_SET  = os.getenv("BRAIN_READ_SET", "legacy").strip().lower()
BRAIN_WRITE_SET = os.getenv("BRAIN_WRITE_SET", "legacy").strip().lower()

if BRAIN_READ_SET not in TABLE_SETS:
    raise RuntimeError(f"BRAIN_READ_SET must be 'legacy' or 'hdi', got {BRAIN_READ_SET!r}")
if BRAIN_WRITE_SET not in ("legacy", "hdi", "both"):
    raise RuntimeError(f"BRAIN_WRITE_SET must be 'legacy', 'hdi' or 'both', got {BRAIN_WRITE_SET!r}")

# Primary first — the set we read from is the one that must succeed.
if BRAIN_WRITE_SET == "both":
    _secondary   = "hdi" if BRAIN_READ_SET == "legacy" else "legacy"
    WRITE_SETS   = [BRAIN_READ_SET, _secondary]
else:
    WRITE_SETS   = [BRAIN_WRITE_SET]

READ_TABLES = TABLE_SETS[BRAIN_READ_SET]
USING_HDI   = BRAIN_READ_SET == "hdi"

# Read-path table names. Every SELECT in this module goes through these.
T_SP_FILES      = READ_TABLES["sp_files"]
T_S3_FILES      = READ_TABLES["s3_files"]
T_TEXT_CHUNKS   = READ_TABLES["text_chunks"]
T_VECTORS       = READ_TABLES["vectors"]
T_FILES_CHUNK   = READ_TABLES["files_chunk"]
T_FILE_REGISTRY = READ_TABLES["file_registry"]
T_SP_VERSIONS   = READ_TABLES["sp_versions"]

log.info(f"BRAIN tables — read={BRAIN_READ_SET} write={'+'.join(WRITE_SETS)} (vectors → {T_VECTORS})")


def _require_legacy_for_ddl(operation: str) -> None:
    """Block DDL that would target the HDI-owned table set."""
    if USING_HDI:
        raise HTTPException(
            status_code=409,
            detail=(
                f"{operation} is not available while BRAIN_READ_SET=hdi. The "
                f"CATALYSTPLATFORM_TB_BRAIN_* tables are owned by the HDI db module — "
                f"change db/src/*.hdbmigrationtable in BTP-Catalyst-Platform and redeploy."
            ),
        )


# Retrieval filter tags. Carried on the chunk and vector tables so a similarity
# search can restrict candidates before ranking.
TAG_COLUMNS = ("layer", "layer_category", "deliverable_code", "project", "client_name")

# Optional tag columns that may or may not exist in legacy tables.
# source comes first because it's inserted as a positional arg before _tags_of().
_OPTIONAL_TAG_COL_NAMES = ("source", "layer", "layer_category", "deliverable_code", "project", "client_name")


def _optional_tag_cols(cur, table: str) -> tuple:
    """Return which optional tag columns actually exist in `table`, in canonical order.

    Cached per process — a restart picks up schema changes. Used to build dynamic
    UPSERTs so ingest works even when the legacy tables have not had all tag columns
    added yet (ALTER TABLE may require DBA privileges the runtime user doesn't have).
    """
    key = table.upper()
    if key not in _col_cache:
        cur.execute(
            "SELECT COLUMN_NAME FROM SYS.TABLE_COLUMNS "
            "WHERE SCHEMA_NAME = CURRENT_SCHEMA AND TABLE_NAME = ? "
            "AND UPPER(COLUMN_NAME) IN ('SOURCE','LAYER','LAYER_CATEGORY','DELIVERABLE_CODE','PROJECT','CLIENT_NAME')",
            (key,)
        )
        _col_cache[key] = frozenset(r[0].lower() for r in cur.fetchall())
    present = _col_cache[key]
    return tuple(c for c in _OPTIONAL_TAG_COL_NAMES if c in present)


def _tag_values(item: dict, parent, source: str, opt_cols: tuple) -> tuple:
    """Return tag values in the same order as opt_cols, for dynamic UPSERTs."""
    tags = _tags_of(item, parent)   # (layer, layer_category, deliverable_code, project, client_name)
    mapping = {
        "source":           source,
        "layer":            tags[0],
        "layer_category":   tags[1],
        "deliverable_code": tags[2],
        "project":          tags[3],
        "client_name":      tags[4],
    }
    return tuple(mapping[c] for c in opt_cols)


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

def _connect_ddl() -> dbapi.Connection:
    """Fresh connection using the HDI deployer user (hdi_user / _DT) which has ALTER TABLE rights.

    Returns a new, uncached connection — the caller must close it. Falls back to
    the runtime user (user / _RT) when hdi_user is absent (local dev with env vars).
    """
    creds = _get_creds()
    ddl_user     = creds.get("hdi_user") or creds.get("user", "")
    ddl_password = creds.get("hdi_password") or creds.get("password", "")
    if not creds.get("host") or not ddl_user:
        raise RuntimeError("HANA DDL credentials not found")
    connect_kwargs = {
        "address":               creds["host"],
        "port":                  int(creds.get("port", 443)),
        "user":                  ddl_user,
        "password":              ddl_password,
        "encrypt":               creds.get("encrypt", True),
        "sslValidateCertificate": creds.get("sslValidateCertificate", True),
    }
    schema = creds.get("schema") or _schema
    if schema:
        connect_kwargs["currentSchema"] = schema
    log.info(f"DDL connection — user={ddl_user} schema={schema}")
    return dbapi.connect(**connect_kwargs)


def _connect() -> dbapi.Connection:
    global _conn, _schema
    with _conn_lock:
        try:
            if _conn is not None and _conn.isconnected():
                return _conn
        except Exception:
            _conn = None
            _col_cache.clear()
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
            log.error(f"HANA connection failed: {e}", exc_info=True)
            raise

        return _conn


# ── Connection test ───────────────────────────────────────────────────────────

def connection_info() -> dict:
    """Test HANA connectivity and return host, user, schema, and DB version."""
    result = {}
    try:
        creds = _get_creds()
        result["host"]   = creds.get("host", "(none)")
        result["port"]   = creds.get("port", "(none)")
        result["user"]   = creds.get("user", "(none)")
        result["schema"] = creds.get("schema", "(none)")
        print(f"[DB] Testing connection — host={result['host']} user={result['user']} schema={result['schema']}")
    except Exception as e:
        result["step_creds"] = f"FAILED: {e}"
        return result

    try:
        conn = _connect()
        result["connected"] = conn.isconnected()
        cur = conn.cursor()
        cur.execute("SELECT VERSION FROM SYS.M_DATABASE")
        row = cur.fetchone()
        result["db_version"] = row[0] if row else "(unknown)"
        cur.close()
        result["step_connect"] = "ok"
        print(f"[DB] Connection ok — user={result['user']} schema={result['schema']} version={result['db_version']}")
    except Exception as e:
        result["step_connect"] = f"FAILED: {e}"
        print(f"[DB] Connection FAILED — {e}")

    return result


# ── Schema ────────────────────────────────────────────────────────────────────
#
# DDL here targets the LEGACY table set only. When BRAIN_READ_SET=hdi the schema
# is owned by BTP-Catalyst-Platform/db/src/*.hdbmigrationtable, and every function
# in this section refuses to run — use describe_tables() to inspect instead.


def create_tables() -> dict:
    """Create the legacy BRAIN_* tables if absent. Idempotent.

    The column list is kept in lockstep with the HDI .hdbmigrationtable files so
    both table sets have the same shape and rows can be copied between them
    without a column-by-column mapping.
    """
    _require_legacy_for_ddl("create_tables")
    conn   = _connect_ddl()
    cur    = conn.cursor()
    status = {}
    try:
        cur.execute("SELECT TABLE_NAME FROM SYS.TABLES WHERE SCHEMA_NAME = CURRENT_SCHEMA")
        existing = {row[0] for row in cur.fetchall()}

        tables = {
            T_SP_FILES: f"""
                CREATE TABLE {_t(T_SP_FILES)} (
                    item_id          NVARCHAR(500) PRIMARY KEY,
                    filename         NVARCHAR(500),
                    file_type        NVARCHAR(200),
                    extension        NVARCHAR(20),
                    size_text        NVARCHAR(50),
                    size_bytes       BIGINT,
                    web_url          NVARCHAR(1000),
                    doc_path         NVARCHAR(1000),
                    last_modified    NVARCHAR(50),
                    created_date     NVARCHAR(50),
                    created_by       NVARCHAR(500),
                    modified_by      NVARCHAR(500),
                    shared           NVARCHAR(100),
                    version_count    INT,
                    source           NVARCHAR(20) DEFAULT 'sharepoint',
                    layer            NVARCHAR(100),
                    layer_category   NVARCHAR(200),
                    deliverable_code NVARCHAR(20),
                    project          NVARCHAR(100),
                    indexed_at       TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    insert_dt        TIMESTAMP
                )
            """,
            T_TEXT_CHUNKS: f"""
                CREATE TABLE {_t(T_TEXT_CHUNKS)} (
                    item_id          NVARCHAR(500),
                    chunk_id         INT,
                    chunk_data       NCLOB,
                    filename         NVARCHAR(500),
                    source           NVARCHAR(20),
                    layer            NVARCHAR(100),
                    layer_category   NVARCHAR(200),
                    deliverable_code NVARCHAR(20),
                    project          NVARCHAR(100),
                    insert_dt        TIMESTAMP,
                    PRIMARY KEY (item_id, chunk_id)
                )
            """,
            T_VECTORS: f"""
                CREATE TABLE {_t(T_VECTORS)} (
                    item_id          NVARCHAR(500),
                    chunk_id         INT,
                    vector_data      REAL_VECTOR(384),
                    filename         NVARCHAR(500),
                    source           NVARCHAR(20),
                    layer            NVARCHAR(100),
                    layer_category   NVARCHAR(200),
                    deliverable_code NVARCHAR(20),
                    project          NVARCHAR(100),
                    insert_dt        TIMESTAMP,
                    PRIMARY KEY (item_id, chunk_id)
                )
            """,
            T_FILE_REGISTRY: f"""
                CREATE TABLE {_t(T_FILE_REGISTRY)} (
                    item_id           NVARCHAR(500) PRIMARY KEY,
                    modified_datetime NVARCHAR(50),
                    source            NVARCHAR(20),
                    filename          NVARCHAR(500),
                    insert_datetime   TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """,
            T_S3_FILES: f"""
                CREATE TABLE {_t(T_S3_FILES)} (
                    item_id          NVARCHAR(500) PRIMARY KEY,
                    s3_key           NVARCHAR(1000),
                    filename         NVARCHAR(500),
                    extension        NVARCHAR(20),
                    size_text        NVARCHAR(50),
                    size_bytes       BIGINT,
                    etag             NVARCHAR(200),
                    storage_class    NVARCHAR(50),
                    last_modified    NVARCHAR(50),
                    source           NVARCHAR(20) DEFAULT 's3',
                    layer            NVARCHAR(100),
                    layer_category   NVARCHAR(200),
                    deliverable_code NVARCHAR(20),
                    project          NVARCHAR(100),
                    indexed_at       TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    insert_dt        TIMESTAMP
                )
            """,
            T_FILES_CHUNK: f"""
                CREATE TABLE {_t(T_FILES_CHUNK)} (
                    item_id        NVARCHAR(500) NOT NULL,
                    chunk_id       INT NOT NULL,
                    filename       NVARCHAR(500),
                    chunk_count    INT,
                    chunk_size     INT,
                    chunk_data     NCLOB,
                    chunk_datetime NVARCHAR(50),
                    source         NVARCHAR(20),
                    insert_dt      TIMESTAMP,
                    PRIMARY KEY (item_id, chunk_id)
                )
            """,
            T_SP_VERSIONS: f"""
                CREATE TABLE {_t(T_SP_VERSIONS)} (
                    item_id          NVARCHAR(500) NOT NULL,
                    version_id       NVARCHAR(100) NOT NULL,
                    file_size_text   NVARCHAR(50),
                    size_bytes       BIGINT,
                    last_modified_dt NVARCHAR(50),
                    lastmodified_by  NVARCHAR(500),
                    download_url     NVARCHAR(2000),
                    version_count    INT,
                    filename         NVARCHAR(500),
                    insert_dt        TIMESTAMP,
                    PRIMARY KEY (item_id, version_id)
                )
            """,
        }

        for table_name, ddl in tables.items():
            if table_name in existing:
                status[table_name] = "already exists"
                log.info(f"{_t(table_name)} — already exists")
            else:
                cur.execute(ddl)
                status[table_name] = "created"
                log.info(f"{_t(table_name)} — created")

        conn.commit()
        log.info(f"HANA tables ready — {status}")
        return status
    finally:
        cur.close()
        conn.close()


def describe_tables() -> dict:
    """Report presence and row counts for BOTH table sets, side by side.

    Creates nothing. This is the check after a dual-write run: the per-key
    "matched" flag says whether legacy and HDI hold the same number of rows, which
    is the thing you actually want to know before trusting the new tables.
    """
    conn = _connect()
    cur  = conn.cursor()
    try:
        cur.execute("SELECT TABLE_NAME FROM SYS.TABLES WHERE SCHEMA_NAME = CURRENT_SCHEMA")
        existing = {row[0] for row in cur.fetchall()}

        def inspect(table: str) -> dict:
            if table not in existing:
                return {"table": table, "present": False, "rows": None}
            try:
                cur.execute(f"SELECT COUNT(*) FROM {_t(table)}")
                return {"table": table, "present": True, "rows": cur.fetchone()[0]}
            except Exception as e:
                return {"table": table, "present": True, "rows": None, "error": str(e)}

        comparison = {}
        for key in LEGACY_TABLES:
            legacy = inspect(LEGACY_TABLES[key])
            hdi    = inspect(HDI_TABLES[key])
            comparison[key] = {
                "legacy":  legacy,
                "hdi":     hdi,
                "matched": legacy["rows"] is not None and legacy["rows"] == hdi["rows"],
            }

        return {
            "read_set":   BRAIN_READ_SET,
            "write_sets": WRITE_SETS,
            "schema":     _schema,
            "tables":     comparison,
        }
    finally:
        cur.close()


# ── Ingest ────────────────────────────────────────────────────────────────────

def _tags_of(record: dict, fallback: dict = None) -> tuple:
    """Pull the four retrieval tags off a chunk/vector/item dict, in column order.

    Chunk and vector records carry their own tags when the caller set them; if not,
    they inherit from the parent file item. Empty strings normalise to None so a
    missing tag never matches an equality filter.
    """
    fallback = fallback or {}
    out = []
    for col in TAG_COLUMNS:
        value = record.get(col)
        if value in (None, ""):
            value = fallback.get(col)
        out.append(value or None)
    return tuple(out)


def _fan_out(operation: str, write_fn) -> dict:
    """Run a write against every configured table set.

    write_fn(tables) is called once per set, primary first. The primary set's
    failure propagates; a secondary failure is captured so a problem with the new
    HDI tables can never take down the live legacy path. The returned dict always
    carries an "errors" key — empty on a clean run, which is what tests assert on.
    """
    result = {"written_to": [], "errors": {}}
    for index, set_name in enumerate(WRITE_SETS):
        tables = TABLE_SETS[set_name]
        try:
            outcome = write_fn(tables)
            result["written_to"].append(set_name)
            if index == 0 and isinstance(outcome, dict):
                result.update(outcome)
            elif isinstance(outcome, dict):
                result[set_name] = outcome
        except Exception as e:
            if index == 0:
                raise
            log.error(
                f"[DB] {operation}: secondary write to '{set_name}' FAILED "
                f"({type(e).__name__}: {e}) — primary set '{WRITE_SETS[0]}' is unaffected",
                exc_info=True,
            )
            result["errors"][set_name] = f"{type(e).__name__}: {e}"
    return result


def ingest(chunks_data: list, vectors_data: list, items: list, source: str = "sharepoint") -> dict:
    """Upsert files, text chunks and vectors into every configured table set.

    Each chunk and vector row carries layer/layer_category/deliverable_code/project
    so POST /db/query can filter the candidate set before the cosine ranking runs.
    Tags come off the record itself, falling back to the parent file item.

    With BRAIN_WRITE_SET=both the same rows land in the legacy and HDI tables, so
    the new schema is exercised by ordinary API traffic. See _fan_out for how a
    failure on the secondary set is contained.
    """
    return _fan_out(
        "ingest",
        lambda tables: _ingest_into(tables, chunks_data, vectors_data, items, source),
    )


def _ingest_into(tables: dict, chunks_data: list, vectors_data: list, items: list, source: str) -> dict:
    T_SP_FILES    = tables["sp_files"]
    T_S3_FILES    = tables["s3_files"]
    T_TEXT_CHUNKS = tables["text_chunks"]
    T_VECTORS     = tables["vectors"]

    conn      = _connect()
    cur       = conn.cursor()
    files_map = {item["id"]: item for item in items if item.get("id")}
    now       = datetime.utcnow()
    # item_id → parent item, so a chunk with no tags of its own inherits the file's.
    tag_source = {item_id: item for item_id, item in files_map.items()}
    try:
        if source == "s3":
            log.info(f"[DB STEP 1/3] Upserting {len(files_map)} file(s) into {T_S3_FILES}")
            for item_id, item in files_map.items():
                log.debug(f"  {T_S3_FILES} upsert — item_id={item_id} filename={item.get('name','')}")
                try:
                    _s3_opt  = _optional_tag_cols(cur, T_S3_FILES)
                    _s3_core = ["item_id", "s3_key", "filename", "extension", "size_text",
                                "size_bytes", "etag", "storage_class", "last_modified"]
                    _s3_all  = _s3_core + list(_s3_opt) + ["insert_dt"]
                    cur.execute(
                        f"UPSERT {_t(T_S3_FILES)} ({', '.join(_s3_all)}) "
                        f"VALUES ({', '.join(['?'] * len(_s3_all))}) WITH PRIMARY KEY",
                        (
                            item_id,
                            item.get("doc_path", ""),
                            item.get("name", ""),
                            item.get("extension", ""),
                            item.get("size", ""),
                            item.get("size_bytes"),
                            item.get("etag", ""),
                            item.get("storage_class", "STANDARD"),
                            item.get("last_modified_date", ""),
                            *_tag_values(item, None, source, _s3_opt),
                            now,
                        )
                    )
                except Exception as e:
                    log.error(f"  {T_S3_FILES} upsert FAILED — item_id={item_id} filename={item.get('name','')}: {e}", exc_info=True)
                    raise
        else:
            log.info(f"[DB STEP 1/3] Upserting {len(files_map)} file(s) into {T_SP_FILES}")
            for item_id, item in files_map.items():
                log.debug(f"  {T_SP_FILES} upsert — item_id={item_id} filename={item.get('name','')}")
                try:
                    _sp_opt  = _optional_tag_cols(cur, T_SP_FILES)
                    _sp_core = ["item_id", "filename", "file_type", "extension", "size_text",
                                "size_bytes", "web_url", "doc_path", "last_modified",
                                "created_date", "created_by", "modified_by", "shared", "version_count"]
                    _sp_all  = _sp_core + list(_sp_opt) + ["insert_dt"]
                    cur.execute(
                        f"UPSERT {_t(T_SP_FILES)} ({', '.join(_sp_all)}) "
                        f"VALUES ({', '.join(['?'] * len(_sp_all))}) WITH PRIMARY KEY",
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
                            *_tag_values(item, None, source, _sp_opt),
                            now,
                        )
                    )
                except Exception as e:
                    log.error(f"  {T_SP_FILES} upsert FAILED — item_id={item_id} filename={item.get('name','')}: {e}", exc_info=True)
                    raise

        log.info(f"[DB STEP 2/3] Upserting {len(chunks_data)} chunk(s) into {T_TEXT_CHUNKS}")
        for c in chunks_data:
            log.debug(f"  {T_TEXT_CHUNKS} upsert — item_id={c['item_id']} chunk_id={c['chunk_id']}")
            try:
                _tc_opt  = _optional_tag_cols(cur, T_TEXT_CHUNKS)
                _tc_core = ["item_id", "chunk_id", "chunk_data", "filename"]
                _tc_all  = _tc_core + list(_tc_opt) + ["insert_dt"]
                cur.execute(
                    f"UPSERT {_t(T_TEXT_CHUNKS)} ({', '.join(_tc_all)}) "
                    f"VALUES ({', '.join(['?'] * len(_tc_all))}) WITH PRIMARY KEY",
                    (
                        c["item_id"], c["chunk_id"], c["chunk_data"], c.get("filename", ""),
                        *_tag_values(c, tag_source.get(c["item_id"]), source, _tc_opt),
                        now,
                    )
                )
            except Exception as e:
                log.error(f"  {T_TEXT_CHUNKS} upsert FAILED — item_id={c['item_id']} chunk_id={c['chunk_id']}: {e}", exc_info=True)
                raise

        log.info(f"[DB STEP 3/3] Upserting {len(vectors_data)} vector(s) into {T_VECTORS}")
        for v in vectors_data:
            log.debug(f"  {T_VECTORS} upsert — item_id={v['item_id']} chunk_id={v['chunk_id']}")
            try:
                vec_str  = "[" + ",".join(f"{x:.8f}" for x in v["vector_data"]) + "]"
                _vt_opt  = _optional_tag_cols(cur, T_VECTORS)
                _vt_core = ["item_id", "chunk_id", "vector_data", "filename"]
                _vt_all  = _vt_core + list(_vt_opt) + ["insert_dt"]
                # vector_data needs the TO_REAL_VECTOR() cast; rebuild placeholders manually.
                _vt_ph   = "?, ?, TO_REAL_VECTOR(?), ?" + (", ?" * len(_vt_opt)) + ", ?"
                cur.execute(
                    f"UPSERT {_t(T_VECTORS)} ({', '.join(_vt_all)}) VALUES ({_vt_ph}) WITH PRIMARY KEY",
                    (
                        v["item_id"], v["chunk_id"], vec_str, v.get("filename", ""),
                        *_tag_values(v, tag_source.get(v["item_id"]), source, _vt_opt),
                        now,
                    )
                )
            except Exception as e:
                log.error(f"  {T_VECTORS} upsert FAILED — item_id={v['item_id']} chunk_id={v['chunk_id']}: {e}", exc_info=True)
                raise

        conn.commit()
        result = {"files": len(files_map), "chunks": len(chunks_data), "vectors": len(vectors_data)}
        log.info(f"[DB] Ingest committed — {result}")
        return result
    except Exception as e:
        log.error(f"[DB] Ingest rolling back — {type(e).__name__}: {e}", exc_info=True)
        conn.rollback()
        raise
    finally:
        cur.close()


def ingest_raw_chunks(raw_chunk_records: list, source: str = "sharepoint") -> dict:
    """Upsert raw binary file chunks (base64-encoded) into every configured set."""
    if not raw_chunk_records:
        return {"raw_chunks": 0, "written_to": [], "errors": {}}
    return _fan_out(
        "ingest_raw_chunks",
        lambda tables: _ingest_raw_chunks_into(tables, raw_chunk_records, source),
    )


def _ingest_raw_chunks_into(tables: dict, raw_chunk_records: list, source: str) -> dict:
    T_FILES_CHUNK = tables["files_chunk"]
    conn = _connect()
    cur  = conn.cursor()
    now  = datetime.utcnow()
    try:
        log.info(f"[DB] Upserting {len(raw_chunk_records)} raw chunk(s) into {T_FILES_CHUNK}")
        for rec in raw_chunk_records:
            try:
                cur.execute(
                    f"UPSERT {_t(T_FILES_CHUNK)} "
                    "(item_id, chunk_id, filename, chunk_size, chunk_data, chunk_datetime, source, insert_dt) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?) WITH PRIMARY KEY",
                    (
                        rec["item_id"],
                        rec["chunk_id"],
                        rec.get("filename", ""),
                        rec.get("chunk_size", 0),
                        rec.get("chunk_data", ""),
                        rec.get("chunk_datetime", ""),
                        source,
                        now,
                    )
                )
            except Exception as e:
                log.error(f"  {T_FILES_CHUNK} upsert FAILED — item_id={rec['item_id']} chunk_id={rec['chunk_id']}: {e}", exc_info=True)
                raise
        conn.commit()
        result = {"raw_chunks": len(raw_chunk_records)}
        log.info(f"[DB] ingest_raw_chunks committed — {result}")
        return result
    except Exception as e:
        log.error(f"[DB] ingest_raw_chunks rolling back — {type(e).__name__}: {e}", exc_info=True)
        conn.rollback()
        raise
    finally:
        cur.close()


def ingest_file_registry(items: list, source: str = "sharepoint") -> dict:
    if not items:
        return {"file_registry": 0, "written_to": [], "errors": {}}
    return _fan_out(
        "ingest_file_registry",
        lambda tables: _ingest_file_registry_into(tables, items, source),
    )


def _ingest_file_registry_into(tables: dict, items: list, source: str) -> dict:
    T_FILE_REGISTRY = tables["file_registry"]
    conn = _connect()
    cur  = conn.cursor()
    now  = datetime.utcnow()
    try:
        for item in items:
            raw_name = item.get("filename") or item.get("key") or item.get("item_id", "")
            filename = raw_name.split("/")[-1] if raw_name else ""
            cur.execute(
                f"UPSERT {_t(T_FILE_REGISTRY)} "
                "(item_id, modified_datetime, source, filename, insert_datetime) "
                "VALUES (?, ?, ?, ?, ?) WITH PRIMARY KEY",
                (
                    item.get("item_id", ""),
                    item.get("modified_datetime") or item.get("last_modified_date", ""),
                    source,
                    filename,
                    now,
                )
            )
        conn.commit()
        log.info(f"Ingested {len(items)} rows into {T_FILE_REGISTRY} (source={source})")
        return {"file_registry": len(items)}
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()



def ingest_versions(versions: list) -> dict:
    """Upsert SharePoint version records into every configured table set."""
    if not versions:
        return {"sp_files_versions": 0, "written_to": [], "errors": {}}
    return _fan_out("ingest_versions", lambda tables: _ingest_versions_into(tables, versions))


def _ingest_versions_into(tables: dict, versions: list) -> dict:
    T_SP_VERSIONS = tables["sp_versions"]
    conn = _connect()
    cur  = conn.cursor()
    now  = datetime.utcnow()
    try:
        for v in versions:
            cur.execute(
                f"UPSERT {_t(T_SP_VERSIONS)} "
                "(item_id, version_id, file_size_text, size_bytes, last_modified_dt, "
                " lastmodified_by, download_url, version_count, filename, insert_dt) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) WITH PRIMARY KEY",
                (
                    v.get("item_id", ""),
                    v.get("version_id", ""),
                    v.get("fileSize"),
                    v.get("size"),
                    v.get("lastModifiedDateTime"),
                    v.get("lastmodified_by"),
                    v.get("download_url"),
                    v.get("version_count"),
                    v.get("filename"),
                    now,
                )
            )
        conn.commit()
        log.info(f"Ingested {len(versions)} rows into {T_SP_VERSIONS}")
        return {"sp_files_versions": len(versions)}
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()


# ── HDI backfill ──────────────────────────────────────────────────────────────

_BACKFILL_KEYS = ["sp_files", "s3_files", "text_chunks", "vectors",
                   "files_chunk", "file_registry", "sp_versions"]

_BACKFILL_PRIMARY_KEYS = {
    "sp_files":      ["ITEM_ID"],
    "s3_files":      ["ITEM_ID"],
    "text_chunks":   ["ITEM_ID", "CHUNK_ID"],
    "vectors":       ["ITEM_ID", "CHUNK_ID"],
    "files_chunk":   ["ITEM_ID", "CHUNK_ID"],
    "file_registry": ["ITEM_ID"],
    "sp_versions":   ["ITEM_ID", "VERSION_ID"],
}


def _columns_of(cur, table: str) -> list:
    """Column names for a table, in positional order. Empty list if absent."""
    cur.execute(
        "SELECT COLUMN_NAME FROM SYS.TABLE_COLUMNS "
        "WHERE SCHEMA_NAME = CURRENT_SCHEMA AND TABLE_NAME = ? "
        "ORDER BY POSITION",
        (table,)
    )
    return [row[0] for row in cur.fetchall()]


def _count_table(cur, table: str):
    try:
        cur.execute(f"SELECT COUNT(*) FROM {_t(table)}")
        return cur.fetchone()[0]
    except Exception:
        return None


def backfill_hdi(table_key: str = None) -> dict:
    """Copy rows missing from HDI tables using INSERT … NOT EXISTS (idempotent).

    Resolves shared columns from SYS.TABLE_COLUMNS at runtime, so it works even
    when the legacy and HDI schemas differ slightly. Columns present only in HDI
    are left NULL; columns present only in legacy are skipped.

    Returns a dict with per-table results: inserted, source_rows, target_rows, skip, error.
    """
    conn = _connect()
    cur  = conn.cursor()
    keys = [table_key] if table_key else _BACKFILL_KEYS
    results = {}
    try:
        for key in keys:
            source = LEGACY_TABLES[key]
            target = HDI_TABLES[key]
            source_cols = _columns_of(cur, source)
            target_cols = _columns_of(cur, target)

            if not source_cols:
                results[key] = {"skip": "legacy table does not exist"}
                continue
            if not target_cols:
                results[key] = {"skip": "HDI table does not exist — deploy the db module first"}
                continue

            shared = [c for c in source_cols if c in set(target_cols)]
            if not shared:
                results[key] = {"skip": "no columns in common"}
                continue

            col_list = ", ".join(f'"{c}"' for c in shared)
            pk = _BACKFILL_PRIMARY_KEYS[key]
            pk_match = " AND ".join(f't."{c}" = s."{c}"' for c in pk)
            sql = (
                f"INSERT INTO {_t(target)} ({col_list}) "
                f"SELECT {col_list} FROM {_t(source)} s "
                f"WHERE NOT EXISTS (SELECT 1 FROM {_t(target)} t WHERE {pk_match})"
            )

            before = _count_table(cur, target)
            try:
                cur.execute(sql)
                conn.commit()
                after = _count_table(cur, target)
                inserted = (after - before) if (after is not None and before is not None) else None
                results[key] = {
                    "source": source, "target": target,
                    "source_rows": _count_table(cur, source),
                    "target_rows": after, "inserted": inserted,
                }
            except Exception as e:
                conn.rollback()
                results[key] = {"source": source, "target": target, "error": f"{type(e).__name__}: {e}"}
    finally:
        cur.close()
    return results


# ── Retag ─────────────────────────────────────────────────────────────────────

_RETAG_VALID_COLS = frozenset(("source", "layer", "layer_category", "deliverable_code", "project", "client_name"))


def _get_matching_item_ids(cur, tables: dict, filter_spec: dict) -> list:
    """Return item_ids from the files tables that match filter_spec."""
    if "item_ids" in filter_spec:
        return list(filter_spec["item_ids"])
    if "sp_path_contains" in filter_spec:
        val = f"%{filter_spec['sp_path_contains']}%"
        cur.execute(f"SELECT item_id FROM {_t(tables['sp_files'])} WHERE doc_path LIKE ?", [val])
        return [r[0] for r in cur.fetchall()]
    if "s3_key_contains" in filter_spec:
        val = f"%{filter_spec['s3_key_contains']}%"
        cur.execute(f"SELECT item_id FROM {_t(tables['s3_files'])} WHERE s3_key LIKE ?", [val])
        return [r[0] for r in cur.fetchall()]
    if "filename_contains" in filter_spec:
        val = f"%{filter_spec['filename_contains']}%"
        ids = set()
        for key in ("sp_files", "s3_files"):
            try:
                cur.execute(f"SELECT item_id FROM {_t(tables[key])} WHERE filename LIKE ?", [val])
                ids.update(r[0] for r in cur.fetchall())
            except Exception:
                pass
        return list(ids)
    raise ValueError(
        f"Unsupported filter: {list(filter_spec.keys())}. "
        "Use: sp_path_contains, s3_key_contains, filename_contains, or item_ids"
    )


def retag_items(filter_spec: dict, tags: dict, table_set_name: str = "hdi") -> dict:
    """UPDATE tag columns on existing chunk/vector records for items matching filter_spec."""
    if table_set_name not in ("legacy", "hdi", "both"):
        raise ValueError(f"table_set must be legacy, hdi, or both — got {table_set_name!r}")
    invalid = set(tags) - _RETAG_VALID_COLS
    if invalid:
        raise ValueError(f"Invalid tag columns: {sorted(invalid)}. Valid: {sorted(_RETAG_VALID_COLS)}")

    tag_pairs = [
        (c, tags[c])
        for c in ("source", "layer", "layer_category", "deliverable_code", "project", "client_name")
        if c in tags
    ]
    sets_to_update = (
        [("legacy", LEGACY_TABLES), ("hdi", HDI_TABLES)]
        if table_set_name == "both"
        else [(table_set_name, TABLE_SETS[table_set_name])]
    )

    conn = _connect()
    cur  = conn.cursor()
    results = {}
    try:
        for set_name, tables in sets_to_update:
            item_ids = _get_matching_item_ids(cur, tables, filter_spec)
            if not item_ids:
                results[set_name] = {"matched_items": 0}
                continue

            def _update(table_name):
                available = _optional_tag_cols(cur, table_name)
                cols = [(c, v) for c, v in tag_pairs if c in available]
                if not cols:
                    return {"skipped": "no matching tag columns in table"}
                set_clause = ", ".join(f"{c} = ?" for c, _ in cols)
                vals = [v for _, v in cols]
                total = 0
                for i in range(0, len(item_ids), 900):
                    batch = item_ids[i:i + 900]
                    ph = ", ".join("?" * len(batch))
                    cur.execute(
                        f"UPDATE {_t(table_name)} SET {set_clause} WHERE item_id IN ({ph})",
                        vals + batch,
                    )
                    total += max(cur.rowcount, 0)
                return {"updated": total}

            set_results = {}
            for table_key in ("sp_files", "s3_files", "text_chunks", "vectors"):
                set_results[table_key] = _update(tables[table_key])
            results[set_name] = {**set_results, "matched_items": len(item_ids)}

        conn.commit()
        return results
    except Exception as e:
        conn.rollback()
        raise
    finally:
        cur.close()


# ── Backup ────────────────────────────────────────────────────────────────────

def backup_to_s3(s3_dir: str, include_vectors: bool = True) -> dict:
    """
    Dump all BRAIN tables to gzip-compressed JSONL and upload to S3 under s3_dir.

    Files written:
      <s3_dir>/brain_file_registry_<ts>.jsonl.gz
      <s3_dir>/brain_sp_files_<ts>.jsonl.gz
      <s3_dir>/brain_s3_files_<ts>.jsonl.gz
      <s3_dir>/brain_text_chunks_<ts>.jsonl.gz
      <s3_dir>/brain_files_chunk_<ts>.jsonl.gz
      <s3_dir>/brain_vectors_<ts>.jsonl.gz   (skipped if include_vectors=False)

    Returns dict with per-table row counts, uploaded S3 keys, and compressed sizes.
    """
    import gzip
    import s3 as s3_client

    s3_dir = s3_dir.rstrip("/")
    ts     = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    conn   = _connect()
    cur    = conn.cursor()
    result = {}

    tables = [
        (T_FILE_REGISTRY, f"brain_file_registry_{ts}.jsonl.gz"),
        (T_SP_FILES,      f"brain_sp_files_{ts}.jsonl.gz"),
        (T_S3_FILES,      f"brain_s3_files_{ts}.jsonl.gz"),
        (T_TEXT_CHUNKS,   f"brain_text_chunks_{ts}.jsonl.gz"),
        (T_FILES_CHUNK,   f"brain_files_chunk_{ts}.jsonl.gz"),
        (T_SP_VERSIONS,   f"brain_sp_files_versions_{ts}.jsonl.gz"),
    ]
    if include_vectors:
        tables.append((T_VECTORS, f"brain_vectors_{ts}.jsonl.gz"))

    # CAP TB_* tables — schema-qualified using HANA_CAP_SCHEMA env var
    cap_schema = os.getenv("HANA_CAP_SCHEMA", "")
    cap_tables = [
        "TB_PHASES",
        "TB_ROLES",
        "TB_USERS",
        "TB_PROJECT_DETAILS",
        "TB_PROJECT_MEMBERS",
        "TB_PROJECT_ENGAGEMENT_OVERVIEW",
        "TB_PROJECT_AUDIT",
        "TB_DELIVERABLES",
        "TB_DELIVERABLE_DEPENDENCIES",
        "TB_PROJECT_DELIVERABLES",
        "TB_DELIVERABLE_REVIEWS",
        "TB_UPLOADED_FILES",
        "TB_PROJECT_GATE_CRITERIA",
        "TB_KNOWLEDGE_BASE",
        "TB_HANDOFF_NOTICES",
        "TB_TEAM_ASKS",
    ]
    if cap_schema:
        for tb in cap_tables:
            qualified = f"{cap_schema}.CATALYSTPLATFORM_{tb}"
            tables.append((qualified, f"{tb.lower()}_{ts}.jsonl.gz"))

    try:
        for table, filename in tables:
            s3_key    = f"{s3_dir}/{filename}"
            qualified = table if "." in table else _t(table)
            try:
                cur.execute(f"SELECT * FROM {qualified}")
                cols  = [d[0].lower() for d in cur.description]
                rows  = cur.fetchall()
                lines = []
                for row in rows:
                    record = {}
                    for col, val in zip(cols, row):
                        if hasattr(val, "isoformat"):
                            record[col] = val.isoformat()
                        else:
                            record[col] = val
                    lines.append(json.dumps(record, ensure_ascii=False))

                jsonl_bytes      = ("\n".join(lines)).encode("utf-8")
                compressed_bytes = gzip.compress(jsonl_bytes, compresslevel=6)
                s3_client.upload_bytes(compressed_bytes, s3_key, content_type="application/gzip")
                result[table] = {
                    "rows":             len(rows),
                    "s3_key":           s3_key,
                    "bytes_raw":        len(jsonl_bytes),
                    "bytes_compressed": len(compressed_bytes),
                }
                log.info(f"Backup: {table} → s3:{s3_key} ({len(rows)} rows, {len(compressed_bytes):,} bytes compressed)")
            except Exception as e:
                result[table] = {"rows": 0, "s3_key": s3_key, "error": str(e)}
                log.error(f"Backup failed for {table}: {e}", exc_info=True)
    finally:
        cur.close()

    return result


# ── Query ─────────────────────────────────────────────────────────────────────

# Columns a caller may filter on. Whitelisted rather than interpolated freely —
# the column name goes into the SQL text, so it must never come from user input.
# `source` lives on the vector table too, so every one of these is a single-table
# predicate and restricts the candidate set before COSINE_SIMILARITY is evaluated.
FILTERABLE = ("layer", "layer_category", "deliverable_code", "project", "source", "client_name")


def _build_filters(filters: dict) -> tuple:
    """Turn a {column: value-or-list} dict into (sql_fragment, params).

    A list value becomes an IN (...) so callers can ask for several layers at once
    (e.g. both guideline layers in one round trip). Unknown columns raise rather
    than being silently dropped — a typo'd filter that quietly matched everything
    would hand back the whole KB, which is the exact failure this replaces.
    """
    if not filters:
        return "", []

    clauses, params = [], []
    for col, value in filters.items():
        if value is None or value == [] or value == "":
            continue
        if col not in FILTERABLE:
            raise ValueError(f"Unknown filter column '{col}'. Allowed: {', '.join(FILTERABLE)}")
        if isinstance(value, (list, tuple, set)):
            values = [v for v in value if v not in (None, "")]
            if not values:
                continue
            placeholders = ", ".join("?" for _ in values)
            clauses.append(f"v.{col} IN ({placeholders})")
            params.extend(values)
        else:
            clauses.append(f"v.{col} = ?")
            params.append(value)

    if not clauses:
        return "", []
    return " WHERE " + " AND ".join(clauses), params


def query(vector: list, top_k: int = 5, filters: dict = None, min_score: float = None) -> list:
    """Cosine similarity search over BTP-Brain chunks, optionally layer-filtered.

    `filters` restricts the candidate set (see FILTERABLE); `min_score` drops hits
    below a cosine floor so weakly-related chunks do not dilute the prompt. Both
    are applied in-engine, so top_k counts rows that already passed them.
    """
    vec_str = "[" + ",".join(f"{x:.8f}" for x in vector) + "]"
    where, where_params = _build_filters(filters)

    # min_score filters the computed score, so it belongs in HAVING-like position;
    # wrapping the ranked set in an outer SELECT keeps it out of the WHERE clause.
    score_filter = ""
    params = [vec_str] + where_params
    if min_score is not None:
        score_filter = " AND score >= ?"
        params.append(float(min_score))

    conn = _connect()
    cur  = conn.cursor()
    try:
        cur.execute(f"""
            SELECT TOP {int(top_k)} * FROM (
                SELECT
                    v.item_id,
                    COALESCE(sp.filename, s3.filename, v.filename) AS filename,
                    v.chunk_id, c.chunk_data,
                    v.layer, v.layer_category, v.deliverable_code, v.project, v.source, v.client_name,
                    COSINE_SIMILARITY(v.vector_data, TO_REAL_VECTOR(?)) AS score
                FROM {_t(T_VECTORS)} v
                JOIN {_t(T_TEXT_CHUNKS)} c ON v.item_id = c.item_id AND v.chunk_id = c.chunk_id
                LEFT JOIN {_t(T_SP_FILES)} sp ON v.item_id = sp.item_id
                LEFT JOIN {_t(T_S3_FILES)} s3 ON v.item_id = s3.item_id
                {where}
            ) WHERE 1 = 1{score_filter}
            ORDER BY score DESC
        """, tuple(params))
        return [
            {
                "item_id":          row[0],
                "filename":         row[1],
                "chunk_id":         row[2],
                "chunk_data":       row[3],
                "layer":            row[4],
                "layer_category":   row[5],
                "deliverable_code": row[6],
                "project":          row[7],
                "source":           row[8],
                "client_name":      row[9],
                "score":            float(row[10]),
            }
            for row in cur.fetchall()
        ]
    finally:
        cur.close()


def fetch_chunks(filters: dict, limit: int = 500) -> list:
    """Return chunks by tag alone, ordered by (item, chunk) — no similarity ranking.

    This is the deterministic half of retrieval. A deliverable's template defines
    the document's structure, so it must be injected whole and in order; leaving it
    to top-k similarity would drop sections that happen to score low against the
    query. Guidelines stay semantic; templates come through here.
    """
    where, params = _build_filters(filters)
    if not where:
        raise ValueError("fetch_chunks requires at least one filter — refusing to return the whole KB.")

    conn = _connect()
    cur  = conn.cursor()
    try:
        cur.execute(f"""
            SELECT TOP {int(limit)}
                v.item_id,
                COALESCE(sp.filename, s3.filename, v.filename) AS filename,
                v.chunk_id, c.chunk_data,
                v.layer, v.layer_category, v.deliverable_code, v.project, v.source, v.client_name
            FROM {_t(T_VECTORS)} v
            JOIN {_t(T_TEXT_CHUNKS)} c ON v.item_id = c.item_id AND v.chunk_id = c.chunk_id
            LEFT JOIN {_t(T_SP_FILES)} sp ON v.item_id = sp.item_id
            LEFT JOIN {_t(T_S3_FILES)} s3 ON v.item_id = s3.item_id
            {where}
            ORDER BY filename, v.chunk_id
        """, tuple(params))
        return [
            {
                "item_id":          row[0],
                "filename":         row[1],
                "chunk_id":         row[2],
                "chunk_data":       row[3],
                "layer":            row[4],
                "layer_category":   row[5],
                "deliverable_code": row[6],
                "project":          row[7],
                "source":           row[8],
                "client_name":      row[9],
                "score":            None,
            }
            for row in cur.fetchall()
        ]
    finally:
        cur.close()


# ── CRUD ──────────────────────────────────────────────────────────────────────

def truncate_tables(table_set: str = None) -> dict:
    """Delete all rows from a table set. Returns per-table counts deleted.

    Allowed against the HDI set — DELETE is DML, not DDL, and clearing the new
    tables between test ingests is the whole point of having them. Defaults to
    every configured write set so a dual-write deployment is cleared consistently;
    pass table_set='hdi' to reset only the new tables and leave legacy untouched.
    """
    targets = [table_set] if table_set else list(WRITE_SETS)
    for name in targets:
        if name not in TABLE_SETS:
            raise ValueError(f"Unknown table set {name!r}. Expected one of: {', '.join(TABLE_SETS)}")

    conn = _connect()
    cur  = conn.cursor()
    result = {}
    try:
        for name in targets:
            per_set = {}
            for table in TABLE_SETS[name].values():
                try:
                    cur.execute(f"SELECT COUNT(*) FROM {_t(table)}")
                    before = cur.fetchone()[0]
                    cur.execute(f"DELETE FROM {_t(table)}")
                    per_set[table] = {"deleted": before}
                except Exception as e:
                    per_set[table] = {"deleted": 0, "error": str(e)}
            result[name] = per_set
        conn.commit()
    finally:
        cur.close()
    return result


def drop_tables() -> dict:
    """Drop the orphaned pre-HDI tables. Legacy table set only.

    This has always targeted a short list of dead tables, not the live BRAIN_*
    set — recreate via /db/init if you drop something you still need.
    """
    _require_legacy_for_ddl("drop_tables")
    conn = _connect()
    cur  = conn.cursor()
    # BRAIN_SP_FILES_VERSIONS used to be on this list, but ingest_versions() writes
    # to it — dropping it would have deleted live version history. Removed.
    tables = [
        "CATALYSTPLATFORM_BRAIN_FILES",
        "CATALYSTPLATFORM_BRAIN_SP_FILES_CHUNK",
    ]
    result = {}
    try:
        for table in tables:
            try:
                cur.execute(f"DROP TABLE {_t(table)}")
                result[table] = "dropped"
            except Exception as e:
                result[table] = f"skipped: {e}"
        conn.commit()
    finally:
        cur.close()
    return result


def migrate_tables() -> dict:
    """
    Apply additive column migrations to the legacy BRAIN_* tables.
    Safe to run multiple times — skips columns that already exist.

    Brings the legacy set up to the same shape as the HDI TB_BRAIN_* tables so
    ingest() can write the retrieval tags to either set and rows can be copied
    across without a column mapping. All additions are nullable, so no existing
    row is rewritten and no data is lost.

    Note the tag columns on TEXT_CHUNKS and VECTORS: without them, a filtered
    /db/query against the legacy set has nothing to filter on.
    """
    _require_legacy_for_ddl("migrate_tables")
    conn   = _connect_ddl()
    cur    = conn.cursor()
    status = {}

    # Columns to add: (table, column, definition)
    migrations = [
        (T_SP_FILES,    "source",           "NVARCHAR(20) DEFAULT 'sharepoint'"),
        (T_SP_FILES,    "layer",            "NVARCHAR(100)"),
        (T_SP_FILES,    "layer_category",   "NVARCHAR(200)"),
        (T_SP_FILES,    "deliverable_code", "NVARCHAR(20)"),
        (T_SP_FILES,    "project",          "NVARCHAR(100)"),

        (T_S3_FILES,    "source",           "NVARCHAR(20) DEFAULT 's3'"),
        (T_S3_FILES,    "layer",            "NVARCHAR(100)"),
        (T_S3_FILES,    "layer_category",   "NVARCHAR(200)"),
        (T_S3_FILES,    "deliverable_code", "NVARCHAR(20)"),
        (T_S3_FILES,    "project",          "NVARCHAR(100)"),

        (T_TEXT_CHUNKS, "source",           "NVARCHAR(20)"),
        (T_TEXT_CHUNKS, "layer",            "NVARCHAR(100)"),
        (T_TEXT_CHUNKS, "layer_category",   "NVARCHAR(200)"),
        (T_TEXT_CHUNKS, "deliverable_code", "NVARCHAR(20)"),
        (T_TEXT_CHUNKS, "project",          "NVARCHAR(100)"),

        (T_VECTORS,     "source",           "NVARCHAR(20)"),
        (T_VECTORS,     "layer",            "NVARCHAR(100)"),
        (T_VECTORS,     "layer_category",   "NVARCHAR(200)"),
        (T_VECTORS,     "deliverable_code", "NVARCHAR(20)"),
        (T_VECTORS,     "project",          "NVARCHAR(100)"),

        (T_FILES_CHUNK, "source",           "NVARCHAR(20)"),
    ]

    try:
        # Fetch existing columns once
        cur.execute(
            "SELECT TABLE_NAME, COLUMN_NAME FROM SYS.TABLE_COLUMNS "
            "WHERE SCHEMA_NAME = CURRENT_SCHEMA"
        )
        existing_cols = {(row[0], row[1]) for row in cur.fetchall()}

        existing_tables = {t for t, _ in existing_cols}

        for table, column, definition in migrations:
            key = (table, column.upper())
            if table not in existing_tables:
                status[f"{table}.{column}"] = "skipped: table does not exist"
                continue
            if key in existing_cols:
                status[f"{table}.{column}"] = "already exists"
                log.info(f"{_t(table)}.{column} — already exists")
                continue
            # One failed ALTER must not roll back the ones that already succeeded.
            try:
                cur.execute(
                    f"ALTER TABLE {_t(table)} ADD ({column} {definition})"
                )
                status[f"{table}.{column}"] = "added"
                log.info(f"{_t(table)}.{column} — added")
            except Exception as e:
                status[f"{table}.{column}"] = f"failed: {e}"
                log.error(f"{_t(table)}.{column} — ALTER failed: {e}")

        conn.commit()
        log.info(f"Migration complete — {status}")
        return status
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()
        conn.close()


def get_table_stats() -> dict:
    """Return row counts for all BRAIN tables, with source breakdown where available."""
    conn = _connect()
    cur  = conn.cursor()
    try:
        # (table_name, key, has_source_col)
        tables = [
            (T_SP_FILES,      "BRAIN_SP_FILES",          True),
            (T_S3_FILES,      "BRAIN_S3_FILES",          True),
            (T_TEXT_CHUNKS,   "BRAIN_TEXT_CHUNKS",       True),
            (T_VECTORS,       "BRAIN_VECTORS",           True),
            (T_FILES_CHUNK,   "BRAIN_FILES_CHUNK",       True),
            (T_FILE_REGISTRY, "BRAIN_FILE_REGISTRY",     True),
            (T_SP_VERSIONS,   "BRAIN_SP_FILES_VERSIONS", False),
        ]
        stats = {}
        for table, key, has_source in tables:
            try:
                cur.execute(f"SELECT COUNT(*) FROM {_t(table)}")
                total = cur.fetchone()[0]
            except Exception:
                stats[key] = {"total": "error", "by_source": {}}
                continue
            by_source = {}
            if has_source:
                try:
                    cur.execute(
                        f"SELECT COALESCE(source, 'unknown'), COUNT(*) "
                        f"FROM {_t(table)} GROUP BY source"
                    )
                    by_source = {row[0]: row[1] for row in cur.fetchall()}
                except Exception:
                    pass
            stats[key] = {"total": total, "by_source": by_source}
        return stats
    finally:
        cur.close()


def get_modified_map() -> dict:
    """Return {item_id: last_modified} for every row in BRAIN_SP_FILES — used for change detection."""
    conn = _connect()
    cur  = conn.cursor()
    try:
        cur.execute(
            f"SELECT item_id, last_modified FROM {_t(T_SP_FILES)}"
        )
        return {row[0]: row[1] for row in cur.fetchall()}
    finally:
        cur.close()


def get_files() -> list:
    conn = _connect()
    cur  = conn.cursor()
    try:
        cur.execute(
            f"SELECT item_id, filename, file_type, extension, size_text, size_bytes, "
            f"web_url, doc_path, last_modified, created_date, created_by, modified_by, "
            f"shared, version_count, indexed_at, insert_dt "
            f"FROM {_t(T_SP_FILES)} ORDER BY indexed_at DESC"
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
            f"FROM {_t(T_TEXT_CHUNKS)} WHERE item_id = ? ORDER BY chunk_id",
            (item_id,)
        )
        cols = [d[0].lower() for d in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()]
    finally:
        cur.close()



def get_db_analysis() -> dict:
    """Collect all metrics for the DB analysis dashboard."""
    conn = _connect()
    cur  = conn.cursor()
    try:
        bf  = _t(T_SP_FILES)
        btc = _t(T_TEXT_CHUNKS)
        bv  = _t(T_VECTORS)
        out = {}

        cur.execute(f"SELECT COUNT(*), SUM(size_bytes) FROM {bf}")
        r = cur.fetchone()
        out["total_files"]      = r[0] or 0
        out["total_size_bytes"] = int(r[1]) if r[1] else 0

        cur.execute(f"SELECT COUNT(*) FROM {btc}")
        out["total_chunks"] = cur.fetchone()[0] or 0

        cur.execute(f"SELECT COUNT(*) FROM {bv}")
        out["total_vectors"] = cur.fetchone()[0] or 0

        cur.execute(f"SELECT AVG(LENGTH(chunk_data)) FROM {btc}")
        r = cur.fetchone()
        out["avg_chunk_chars"] = int(r[0]) if r[0] else 0

        cur.execute(
            f"SELECT COALESCE(extension,'unknown'), COUNT(*), SUM(size_bytes) "
            f"FROM {bf} GROUP BY extension ORDER BY COUNT(*) DESC"
        )
        out["by_extension"] = [
            {"ext": r[0], "count": r[1], "size_bytes": int(r[2]) if r[2] else 0}
            for r in cur.fetchall()
        ]

        cur.execute(f"SELECT COALESCE(TO_VARCHAR(shared),'unknown'), COUNT(*) FROM {bf} GROUP BY shared")
        out["shared_breakdown"] = {r[0]: r[1] for r in cur.fetchall()}


        cur.execute(
            f"SELECT TO_VARCHAR(insert_dt,'YYYY-MM-DD'), COUNT(*) "
            f"FROM {bf} GROUP BY TO_VARCHAR(insert_dt,'YYYY-MM-DD') ORDER BY 1"
        )
        out["indexed_by_day"] = [{"day": str(r[0]), "count": r[1]} for r in cur.fetchall()]

        cur.execute(
            f"SELECT filename, COALESCE(extension,'—'), COALESCE(size_text,'—'), "
            f"TO_VARCHAR(insert_dt,'YYYY-MM-DD HH24:MI') "
            f"FROM {bf} ORDER BY insert_dt DESC LIMIT 10"
        )
        out["recent_files"] = [
            {"filename": r[0], "extension": r[1], "size_text": r[2], "insert_dt": str(r[3])}
            for r in cur.fetchall()
        ]

        cur.execute(
            f"SELECT COUNT(*) FROM {btc} c "
            f"WHERE NOT EXISTS "
            f"(SELECT 1 FROM {bv} v WHERE v.item_id = c.item_id AND v.chunk_id = c.chunk_id)"
        )
        out["chunks_without_vectors"] = cur.fetchone()[0] or 0

        cur.execute(
            f"SELECT COALESCE(modified_by,'unknown'), COUNT(*) "
            f"FROM {bf} GROUP BY modified_by ORDER BY COUNT(*) DESC LIMIT 5"
        )
        out["top_modifiers"] = [{"name": r[0], "count": r[1]} for r in cur.fetchall()]

        cur.execute(
            f"SELECT COALESCE(created_by,'unknown'), COUNT(*) "
            f"FROM {bf} GROUP BY created_by ORDER BY COUNT(*) DESC LIMIT 5"
        )
        out["top_creators"] = [{"name": r[0], "count": r[1]} for r in cur.fetchall()]

        # ── Table stats ───────────────────────────────────────────────────────
        out["table_stats"] = get_table_stats()

        # ── S3 metrics ────────────────────────────────────────────────────────
        bs3 = _t(T_S3_FILES)

        cur.execute(f"SELECT COUNT(*), SUM(size_bytes) FROM {bs3}")
        r = cur.fetchone()
        out["s3_total_files"]      = r[0] or 0
        out["s3_total_size_bytes"] = int(r[1]) if r[1] else 0

        cur.execute(
            f"SELECT COALESCE(extension,'unknown'), COUNT(*), SUM(size_bytes) "
            f"FROM {bs3} GROUP BY extension ORDER BY COUNT(*) DESC"
        )
        out["s3_by_extension"] = [
            {"ext": r[0], "count": r[1], "size_bytes": int(r[2]) if r[2] else 0}
            for r in cur.fetchall()
        ]

        cur.execute(
            f"SELECT TO_VARCHAR(insert_dt,'YYYY-MM-DD'), COUNT(*) "
            f"FROM {bs3} GROUP BY TO_VARCHAR(insert_dt,'YYYY-MM-DD') ORDER BY 1"
        )
        out["s3_indexed_by_day"] = [{"day": str(r[0]), "count": r[1]} for r in cur.fetchall()]

        cur.execute(
            f"SELECT filename, COALESCE(extension,'—'), COALESCE(size_text,'—'), "
            f"COALESCE(storage_class,'—'), TO_VARCHAR(insert_dt,'YYYY-MM-DD HH24:MI') "
            f"FROM {bs3} ORDER BY insert_dt DESC LIMIT 10"
        )
        out["s3_recent_files"] = [
            {"filename": r[0], "extension": r[1], "size_text": r[2], "storage_class": r[3], "insert_dt": str(r[4])}
            for r in cur.fetchall()
        ]

        return out
    finally:
        cur.close()


def delete_item(item_id: str) -> dict:
    """Remove an item from every configured table set.

    Fans out for the same reason ingest does: with dual-write on, deleting from
    only one set would silently leave the other holding rows the first no longer
    has, and the row-count comparison in describe_tables() would start lying.
    """
    return _fan_out("delete_item", lambda tables: _delete_item_from(tables, item_id))


def _delete_item_from(tables: dict, item_id: str) -> dict:
    conn = _connect()
    cur  = conn.cursor()
    try:
        for key in ("vectors", "text_chunks", "sp_files", "s3_files", "files_chunk", "file_registry"):
            cur.execute(f"DELETE FROM {_t(tables[key])} WHERE item_id = ?", (item_id,))
        conn.commit()
        log.info(f"Deleted item {item_id} from {tables['vectors']} and siblings")
        return {"deleted": item_id}
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()


# ── Router ────────────────────────────────────────────────────────────────────

def _require_destructive():
    if os.getenv("ENABLE_DESTRUCTIVE", "false").lower() != "true":
        raise HTTPException(
            status_code=403,
            detail="Destructive actions are disabled on this instance. "
                   "Set ENABLE_DESTRUCTIVE=true and restart the app to enable.",
        )


router = APIRouter(prefix="/db", tags=["database"])


class QueryRequest(BaseModel):
    text:  str = Field(...,  description="Natural language query to search for")
    top_k: int = Field(5, ge=1, le=200, description="Number of results to return")

    # Retrieval filters. Each accepts a single value or a list (list → IN). Omit
    # or null to leave that dimension unfiltered. Filters are applied before the
    # cosine ranking, so top_k counts rows that already matched.
    layer:            str | list[str] | None = Field(
        None, description="BTP-Brain layer, e.g. 'Layer 1 - SAP Guidelines'. List for multiple layers.")
    layer_category:   str | list[str] | None = Field(
        None, description="Sub-folder within the layer.")
    deliverable_code: str | list[str] | None = Field(
        None, description="Deliverable rowId the chunk belongs to, e.g. 'B10'.")
    project:          str | list[str] | None = Field(
        None, description="Project the chunk is scoped to. Null for global reference material.")
    source:           str | list[str] | None = Field(
        None, description="Ingest source: 'sharepoint' or 's3'.")

    min_score: float | None = Field(
        None, ge=-1.0, le=1.0,
        description="Cosine score floor. Chunks below this are dropped so weak matches do not dilute the prompt.")

    def as_filters(self) -> dict:
        return {
            "layer":            self.layer,
            "layer_category":   self.layer_category,
            "deliverable_code": self.deliverable_code,
            "project":          self.project,
            "source":           self.source,
        }


class ChunkFetchRequest(BaseModel):
    """Deterministic tag-only fetch — no query text, no similarity ranking.

    Used to pull a deliverable's template in full, in document order.
    """
    layer:            str | list[str] | None = Field(None, description="BTP-Brain layer.")
    layer_category:   str | list[str] | None = Field(None, description="Sub-folder within the layer.")
    deliverable_code: str | list[str] | None = Field(None, description="Deliverable rowId, e.g. 'B10'.")
    project:          str | list[str] | None = Field(None, description="Project scope.")
    source:           str | list[str] | None = Field(None, description="Ingest source.")
    limit:            int = Field(500, ge=1, le=5000, description="Max chunks returned.")

    def as_filters(self) -> dict:
        return {
            "layer":            self.layer,
            "layer_category":   self.layer_category,
            "deliverable_code": self.deliverable_code,
            "project":          self.project,
            "source":           self.source,
        }


class SqlRequest(BaseModel):
    sql:    str        = Field(..., description="SELECT-only SQL query. Use ? placeholders for values.")
    params: list       = Field(default_factory=list, description="Positional parameters for ? placeholders")
    limit:  int        = Field(100, ge=1, le=5000, description="Max rows returned")


def _validate_sql(sql: str) -> None:
    """
    Rejects anything that is not a plain SELECT:
      - non-SELECT first keyword
      - DML / DDL / procedural keywords
      - comment sequences (-- and /* */)
      - semicolons (statement stacking)
    Values must be passed via the params list, not embedded as literals,
    to prevent injection through user-controlled string values.
    """
    import re
    cleaned = sql.strip()

    # Block SQL comments (-- and /* */)
    if re.search(r"--|/\*|\*/", cleaned):
        raise HTTPException(status_code=400, detail="SQL comments are not permitted")

    # Block semicolons — prevents statement stacking
    if ";" in cleaned:
        raise HTTPException(status_code=400, detail="Semicolons are not permitted")

    # Must begin with SELECT
    first_word = re.split(r"\s+", cleaned, maxsplit=1)[0].upper()
    if first_word != "SELECT":
        raise HTTPException(
            status_code=400,
            detail=f"Only SELECT statements are permitted (received: {first_word})"
        )

    # Block dangerous keywords (word-boundary match, case-insensitive)
    _BLOCKED = re.compile(
        r"\b(INSERT|UPDATE|DELETE|DROP|CREATE|ALTER|TRUNCATE|EXEC|EXECUTE"
        r"|GRANT|REVOKE|MERGE|CALL|PRAGMA|ATTACH|DETACH|INTO|OUTFILE|LOAD)\b",
        re.IGNORECASE,
    )
    hit = _BLOCKED.search(cleaned)
    if hit:
        raise HTTPException(
            status_code=400,
            detail=f"Keyword '{hit.group().upper()}' is not permitted in SQL queries"
        )


@router.post("/query")
def db_query(req: QueryRequest):
    """Layer-filtered semantic search over the BTP-Brain chunks.

    Embeds the query with the same all-MiniLM-L6-v2 model used at vectorize time,
    then ranks by cosine similarity within whatever the filters allow through.
    """
    from api import _get_embed_model
    filters = req.as_filters()
    try:
        vector  = _get_embed_model().encode(req.text).tolist()
        results = query(vector, top_k=req.top_k, filters=filters, min_score=req.min_score)
        return {
            "query":     req.text,
            "filters":   {k: v for k, v in filters.items() if v not in (None, "", [])},
            "min_score": req.min_score,
            "table_set": BRAIN_READ_SET,
            "count":     len(results),
            "results":   results,
        }
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        log.error(f"Query failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/chunks", summary="Fetch chunks by tag, in document order (no ranking)")
def db_fetch_chunks(req: ChunkFetchRequest):
    """Deterministic counterpart to /db/query.

    Returns every chunk matching the tags, ordered by filename then chunk_id, so a
    template comes back whole and in sequence rather than as similarity-ranked
    fragments. At least one filter is required.
    """
    filters = req.as_filters()
    try:
        results = fetch_chunks(filters, limit=req.limit)
        return {
            "filters":   {k: v for k, v in filters.items() if v not in (None, "", [])},
            "table_set": BRAIN_READ_SET,
            "count":     len(results),
            "results":   results,
        }
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        log.error(f"Chunk fetch failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/sql", summary="Run a validated SELECT query against HANA")
def db_sql(req: SqlRequest):
    _validate_sql(req.sql)
    conn = _connect()
    cur  = conn.cursor()
    try:
        cur.execute(req.sql, req.params or [])
        cols = [d[0] for d in (cur.description or [])]
        rows = cur.fetchmany(req.limit)
        result_rows = []
        for row in rows:
            record = {}
            for col, val in zip(cols, row):
                if hasattr(val, "read"):
                    val = val.read()
                if isinstance(val, (bytes, bytearray)):
                    val = val.decode("utf-8", errors="replace")
                record[col] = val
            result_rows.append(record)
        return {"count": len(result_rows), "columns": cols, "rows": result_rows}
    except HTTPException:
        raise
    except Exception as e:
        log.error(f"SQL query failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        cur.close()


@router.get("/files")
def db_files():
    try:
        files = get_files()
        return {"count": len(files), "files": files}
    except Exception as e:
        log.error(f"DB files failed: {type(e).__name__}: {e}", exc_info=True)
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
        log.error(f"DB chunks failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))




@router.get("/vectors/{item_id}")
def db_vectors(item_id: str):
    conn = _connect()
    cur  = conn.cursor()
    try:
        cur.execute(
            f'SELECT chunk_id, vector_data FROM {_t(T_VECTORS)}'
            f' WHERE item_id=? ORDER BY chunk_id',
            (item_id,)
        )
        rows = cur.fetchall()
        if not rows:
            raise HTTPException(status_code=404, detail=f"No vectors found for item_id: {item_id}")
        import struct as _struct, json as _json
        out = []
        for cid, vd in rows:
            if hasattr(vd, "read"):
                vd = vd.read()
            if isinstance(vd, (bytes, bytearray)):
                # HANA REAL_VECTOR binary: [type:uint16 LE][dims:uint16 LE][float32*dims LE]
                num_dims = _struct.unpack_from("<H", vd, 2)[0]
                floats   = list(_struct.unpack_from(f"<{num_dims}f", vd, 4))
            elif isinstance(vd, str):
                floats = _json.loads(vd)
            else:
                floats = list(vd)
            out.append({"chunk_id": cid, "vector": floats})
        return {"item_id": item_id, "count": len(out), "vectors": out}
    except HTTPException:
        raise
    except Exception as e:
        log.error(f"DB vectors failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        cur.close()


@router.delete("/item/{item_id}", dependencies=[Depends(_require_destructive)])
def db_delete(item_id: str):
    try:
        return delete_item(item_id)
    except Exception as e:
        log.error(f"DB delete failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/tables", summary="Which BRAIN table set is active, and are the tables there")
def db_tables():
    """Report the active table set and per-table presence and row counts.

    Read-only — creates nothing. Use this to confirm an HDI deploy landed before
    pointing reads at it with BRAIN_READ_SET=hdi.
    """
    try:
        return describe_tables()
    except Exception as e:
        log.error(f"DB tables failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/init")
def db_init():
    """Create the legacy BRAIN_* tables if absent. Safe to run multiple times.

    Returns 409 when BRAIN_READ_SET=hdi — those tables belong to the HDI db
    module and must be created by deploying it, not from here.
    """
    try:
        result = create_tables()
        return {"status": "ok", "table_set": BRAIN_READ_SET, "tables": result}
    except HTTPException:
        raise
    except Exception as e:
        log.error(f"DB init failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/migrate")
def db_migrate():
    """
    Apply additive column migrations to the legacy BRAIN_* tables (idempotent).
    Adds the source column plus the layer/layer_category/deliverable_code/project
    retrieval tags to SP_FILES, S3_FILES, TEXT_CHUNKS and VECTORS.

    Returns 409 when BRAIN_READ_SET=hdi.
    """
    try:
        result = migrate_tables()
        return {"status": "ok", "table_set": BRAIN_READ_SET, "migrations": result}
    except HTTPException:
        raise
    except Exception as e:
        log.error(f"DB migration failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/drop", dependencies=[Depends(_require_destructive)])
def db_drop(confirm: bool = False):
    """Drop all BRAIN tables entirely. Requires ENABLE_DESTRUCTIVE=true and ?confirm=true."""
    if not confirm:
        raise HTTPException(
            status_code=400,
            detail="Pass ?confirm=true to confirm this destructive action.",
        )
    try:
        result = drop_tables()
        return {"status": "ok", "dropped": result}
    except Exception as e:
        log.error(f"DB drop failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/truncate", dependencies=[Depends(_require_destructive)])
def db_truncate(confirm: bool = False, table_set: str = None):
    """Delete all rows from the BRAIN tables. Requires ENABLE_DESTRUCTIVE=true and ?confirm=true.

    Clears every configured write set by default. Pass ?table_set=hdi to reset only
    the HDI duplicates between test runs, leaving the legacy tables alone.
    """
    if not confirm:
        raise HTTPException(
            status_code=400,
            detail="Pass ?confirm=true to confirm this destructive action.",
        )
    try:
        result = truncate_tables(table_set)
        return {"status": "ok", "deleted": result}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        log.error(f"DB truncate failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/file-registry")
def db_file_registry(source: str = None, limit: int = 200, offset: int = 0):
    conn = _connect()
    cur  = conn.cursor()
    try:
        if source:
            cur.execute(
                f"SELECT item_id, modified_datetime, source, filename, insert_datetime "
                f"FROM {_t(T_FILE_REGISTRY)} "
                f"WHERE source = ? ORDER BY insert_datetime DESC LIMIT ? OFFSET ?",
                (source, limit, offset)
            )
        else:
            cur.execute(
                f"SELECT item_id, modified_datetime, source, filename, insert_datetime "
                f"FROM {_t(T_FILE_REGISTRY)} "
                f"ORDER BY insert_datetime DESC LIMIT ? OFFSET ?",
                (limit, offset)
            )
        cols = [d[0].lower() for d in cur.description]
        rows = [dict(zip(cols, row)) for row in cur.fetchall()]
        return {"count": len(rows), "offset": offset, "rows": rows}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        cur.close()


class BackupRequest(BaseModel):
    s3_dir:          str  = Field(
        ...,
        description=(
            "S3 prefix (folder) where backup files will be stored. "
            "Example: 'backups/2026-09-07' or 'brain/snapshots/prod'. "
            "No trailing slash needed."
        ),
    )
    include_vectors: bool = Field(
        True,
        description="Include BRAIN_VECTORS table in the backup. Vectors can be large — set false to skip.",
    )


@router.post("/backup", summary="Backup all BRAIN tables to S3")
def db_backup(req: BackupRequest):
    """
    Dumps all BRAIN tables to JSONL files and uploads them to the configured S3 bucket
    under the given `s3_dir` prefix.

    **Files written to S3:**
    | File | Table |
    |---|---|
    | `brain_file_registry.jsonl` | BRAIN_FILE_REGISTRY |
    | `brain_sp_files.jsonl` | BRAIN_SP_FILES |
    | `brain_s3_files.jsonl` | BRAIN_S3_FILES |
    | `brain_text_chunks.jsonl` | BRAIN_TEXT_CHUNKS |
    | `brain_vectors.jsonl` | BRAIN_VECTORS (if include_vectors=true) |

    **Example request:**
    ```json
    { "s3_dir": "backups/2026-09-07", "include_vectors": false }
    ```
    """
    try:
        result = backup_to_s3(req.s3_dir, include_vectors=req.include_vectors)
        total_rows = sum(v.get("rows", 0) for v in result.values())
        failed     = [t for t, v in result.items() if "error" in v]
        return {
            "status":     "ok" if not failed else "partial",
            "s3_dir":     req.s3_dir,
            "total_rows": total_rows,
            "failed":     failed,
            "tables":     result,
        }
    except Exception as e:
        log.error(f"DB backup failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


class BackfillRequest(BaseModel):
    table: str = Field(
        None,
        description="Restrict to one logical table key (sp_files, s3_files, text_chunks, "
                    "vectors, files_chunk, file_registry, sp_versions). Omit to copy all.",
    )


@router.post("/backfill", summary="Copy legacy BRAIN_* rows into HDI TB_BRAIN_* tables (idempotent)")
def db_backfill(req: BackfillRequest = BackfillRequest()):
    """Copy rows that exist in the legacy BRAIN_* tables but are missing from the HDI
    TB_BRAIN_* tables, using INSERT … WHERE NOT EXISTS on the primary key.

    Safe to run multiple times — already-copied rows are skipped. Columns that exist
    in legacy only are omitted; columns that exist in HDI only are left NULL.

    Run `GET /db/tables` before and after to verify row counts align.

    **Requires the HDI db module to be deployed first** — if the TB_BRAIN_* tables
    do not exist yet the affected tables will be reported as skipped.

    **Example request (all tables):**
    ```json
    {}
    ```

    **Example request (one table):**
    ```json
    { "table": "vectors" }
    ```
    """
    valid_keys = set(_BACKFILL_KEYS)
    if req.table and req.table not in valid_keys:
        raise HTTPException(status_code=400, detail=f"Unknown table key {req.table!r}. "
                            f"Valid: {sorted(valid_keys)}")
    try:
        results = backfill_hdi(req.table)
        errors  = {k: v["error"] for k, v in results.items() if "error" in v}
        skipped = {k: v["skip"]  for k, v in results.items() if "skip" in v}
        inserted_total = sum(v.get("inserted") or 0 for v in results.values() if "inserted" in v)
        return {
            "status":         "ok" if not errors else "partial",
            "inserted_total": inserted_total,
            "errors":         errors,
            "skipped":        skipped,
            "tables":         results,
        }
    except Exception as e:
        log.error(f"DB backfill failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


class RetagRequest(BaseModel):
    filter: dict = Field(
        ...,
        description=(
            "Exactly one key: sp_path_contains (match doc_path in SP_FILES), "
            "s3_key_contains (match s3_key in S3_FILES), "
            "filename_contains (match filename in both), "
            "or item_ids (explicit list of strings)."
        ),
        examples=[{"sp_path_contains": "BTP-Brain/Layer2/Client Data Sample (Cleaned)/DDA"}],
    )
    tags: dict = Field(
        ...,
        description="Tag columns to set. Any subset of: source, layer, layer_category, deliverable_code, project. Pass null to clear a value.",
        examples=[{"layer": "Layer 2 - Accenture Guidelines", "layer_category": "Client Data Sample", "deliverable_code": "A1"}],
    )
    table_set: str = Field("hdi", description="Which table set to update: hdi (default), legacy, or both.")


@router.post("/retag", summary="UPDATE tag columns on existing chunk/vector records")
def db_retag(req: RetagRequest):
    """Bulk-update layer / layer_category / deliverable_code / project on existing rows.

    Useful when content was vectorized before the tag columns were populated —
    run once per subfolder to label chunks so filtered retrieval can find them.

    **Example — tag all DDA sample chunks:**
    ```json
    {
      "filter": { "sp_path_contains": "Client Data Sample (Cleaned)/DDA" },
      "tags": {
        "layer": "Layer 2 - Accenture Guidelines",
        "layer_category": "Client Data Sample",
        "deliverable_code": "A1"
      },
      "table_set": "hdi"
    }
    ```
    """
    try:
        results = retag_items(req.filter, req.tags, req.table_set)
        total_updated = sum(
            v.get("updated", 0)
            for set_results in results.values()
            if isinstance(set_results, dict)
            for v in set_results.values()
            if isinstance(v, dict)
        )
        return {"status": "ok", "total_updated": total_updated, "sets": results}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        log.error(f"DB retag failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))
