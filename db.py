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
            "CATALYSTPLATFORM_BRAIN_SP_FILES": f"""
                CREATE TABLE {_t("CATALYSTPLATFORM_BRAIN_SP_FILES")} (
                    item_id       NVARCHAR(500) PRIMARY KEY,
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
                    source        NVARCHAR(20) DEFAULT 'sharepoint',
                    indexed_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    insert_dt     TIMESTAMP
                )
            """,
            "CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS": f"""
                CREATE TABLE {_t("CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS")} (
                    item_id    NVARCHAR(500),
                    chunk_id   INT,
                    chunk_data NCLOB,
                    filename   NVARCHAR(500),
                    source     NVARCHAR(20) DEFAULT 'sharepoint',
                    insert_dt  TIMESTAMP,
                    PRIMARY KEY (item_id, chunk_id)
                )
            """,
            "CATALYSTPLATFORM_BRAIN_VECTORS": f"""
                CREATE TABLE {_t("CATALYSTPLATFORM_BRAIN_VECTORS")} (
                    item_id     NVARCHAR(500),
                    chunk_id    INT,
                    vector_data REAL_VECTOR(384),
                    filename    NVARCHAR(500),
                    source      NVARCHAR(20) DEFAULT 'sharepoint',
                    insert_dt   TIMESTAMP,
                    PRIMARY KEY (item_id, chunk_id)
                )
            """,
            "CATALYSTPLATFORM_BRAIN_FILE_REGISTRY": f"""
                CREATE TABLE {_t("CATALYSTPLATFORM_BRAIN_FILE_REGISTRY")} (
                    item_id           NVARCHAR(500) PRIMARY KEY,
                    modified_datetime NVARCHAR(50),
                    source            NVARCHAR(20),
                    filename          NVARCHAR(500),
                    insert_datetime   TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """,
            "CATALYSTPLATFORM_BRAIN_S3_FILES": f"""
                CREATE TABLE {_t("CATALYSTPLATFORM_BRAIN_S3_FILES")} (
                    item_id       NVARCHAR(500) PRIMARY KEY,
                    s3_key        NVARCHAR(1000),
                    filename      NVARCHAR(500),
                    extension     NVARCHAR(20),
                    size_text     NVARCHAR(50),
                    size_bytes    BIGINT,
                    etag          NVARCHAR(200),
                    storage_class NVARCHAR(50),
                    last_modified NVARCHAR(50),
                    indexed_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    insert_dt     TIMESTAMP
                )
            """,
            "CATALYSTPLATFORM_BRAIN_FILES_CHUNK": f"""
                CREATE TABLE {_t("CATALYSTPLATFORM_BRAIN_FILES_CHUNK")} (
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

def ingest(chunks_data: list, vectors_data: list, items: list, source: str = "sharepoint") -> dict:
    conn      = _connect()
    cur       = conn.cursor()
    files_map = {item["id"]: item for item in items if item.get("id")}
    now       = datetime.utcnow()
    try:
        if source == "s3":
            log.info(f"[DB STEP 1/3] Upserting {len(files_map)} file(s) into BRAIN_S3_FILES")
            for item_id, item in files_map.items():
                log.debug(f"  BRAIN_S3_FILES upsert — item_id={item_id} filename={item.get('name','')}")
                try:
                    cur.execute(
                        f"UPSERT {_t('CATALYSTPLATFORM_BRAIN_S3_FILES')} "
                        "(item_id, s3_key, filename, extension, size_text, size_bytes, etag, storage_class, last_modified, insert_dt) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) WITH PRIMARY KEY",
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
                            now,
                        )
                    )
                except Exception as e:
                    log.error(f"  BRAIN_S3_FILES upsert FAILED — item_id={item_id} filename={item.get('name','')}: {e}", exc_info=True)
                    raise
        else:
            log.info(f"[DB STEP 1/3] Upserting {len(files_map)} file(s) into BRAIN_SP_FILES")
            for item_id, item in files_map.items():
                log.debug(f"  BRAIN_SP_FILES upsert — item_id={item_id} filename={item.get('name','')}")
                try:
                    cur.execute(
                        f"UPSERT {_t('CATALYSTPLATFORM_BRAIN_SP_FILES')} "
                        "(item_id, filename, file_type, extension, size_text, size_bytes, web_url, doc_path, "
                        " last_modified, created_date, created_by, modified_by, shared, version_count, source, insert_dt) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) WITH PRIMARY KEY",
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
                            source,
                            now,
                        )
                    )
                except Exception as e:
                    log.error(f"  BRAIN_SP_FILES upsert FAILED — item_id={item_id} filename={item.get('name','')}: {e}", exc_info=True)
                    raise

        log.info(f"[DB STEP 2/3] Upserting {len(chunks_data)} chunk(s) into BRAIN_TEXT_CHUNKS")
        for c in chunks_data:
            log.debug(f"  BRAIN_TEXT_CHUNKS upsert — item_id={c['item_id']} chunk_id={c['chunk_id']}")
            try:
                cur.execute(
                    f"UPSERT {_t('CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS')} "
                    "(item_id, chunk_id, chunk_data, filename, source, insert_dt) "
                    "VALUES (?, ?, ?, ?, ?, ?) WITH PRIMARY KEY",
                    (c["item_id"], c["chunk_id"], c["chunk_data"], c.get("filename", ""), source, now)
                )
            except Exception as e:
                log.error(f"  BRAIN_TEXT_CHUNKS upsert FAILED — item_id={c['item_id']} chunk_id={c['chunk_id']}: {e}", exc_info=True)
                raise

        log.info(f"[DB STEP 3/3] Upserting {len(vectors_data)} vector(s) into BRAIN_VECTORS")
        for v in vectors_data:
            log.debug(f"  BRAIN_VECTORS upsert — item_id={v['item_id']} chunk_id={v['chunk_id']}")
            try:
                vec_str = "[" + ",".join(f"{x:.8f}" for x in v["vector_data"]) + "]"
                cur.execute(
                    f"UPSERT {_t('CATALYSTPLATFORM_BRAIN_VECTORS')} "
                    "(item_id, chunk_id, vector_data, filename, source, insert_dt) "
                    "VALUES (?, ?, TO_REAL_VECTOR(?), ?, ?, ?) WITH PRIMARY KEY",
                    (v["item_id"], v["chunk_id"], vec_str, v.get("filename", ""), source, now)
                )
            except Exception as e:
                log.error(f"  BRAIN_VECTORS upsert FAILED — item_id={v['item_id']} chunk_id={v['chunk_id']}: {e}", exc_info=True)
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
    """Upsert raw binary file chunks (base64-encoded) into BRAIN_FILES_CHUNK."""
    if not raw_chunk_records:
        return {"raw_chunks": 0}
    conn = _connect()
    cur  = conn.cursor()
    now  = datetime.utcnow()
    try:
        log.info(f"[DB] Upserting {len(raw_chunk_records)} raw chunk(s) into BRAIN_FILES_CHUNK")
        for rec in raw_chunk_records:
            try:
                cur.execute(
                    f"UPSERT {_t('CATALYSTPLATFORM_BRAIN_FILES_CHUNK')} "
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
                log.error(f"  BRAIN_FILES_CHUNK upsert FAILED — item_id={rec['item_id']} chunk_id={rec['chunk_id']}: {e}", exc_info=True)
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
        return {"file_registry": 0}
    conn = _connect()
    cur  = conn.cursor()
    now  = datetime.utcnow()
    try:
        for item in items:
            raw_name = item.get("filename") or item.get("key") or item.get("item_id", "")
            filename = raw_name.split("/")[-1] if raw_name else ""
            cur.execute(
                f"UPSERT {_t('CATALYSTPLATFORM_BRAIN_FILE_REGISTRY')} "
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
        log.info(f"Ingested {len(items)} rows into FILE_REGISTRY (source={source})")
        return {"file_registry": len(items)}
    except Exception:
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
        ("CATALYSTPLATFORM_BRAIN_FILE_REGISTRY", f"brain_file_registry_{ts}.jsonl.gz"),
        ("CATALYSTPLATFORM_BRAIN_SP_FILES",      f"brain_sp_files_{ts}.jsonl.gz"),
        ("CATALYSTPLATFORM_BRAIN_S3_FILES",      f"brain_s3_files_{ts}.jsonl.gz"),
        ("CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS",   f"brain_text_chunks_{ts}.jsonl.gz"),
        ("CATALYSTPLATFORM_BRAIN_FILES_CHUNK",   f"brain_files_chunk_{ts}.jsonl.gz"),
    ]
    if include_vectors:
        tables.append(("CATALYSTPLATFORM_BRAIN_VECTORS", f"brain_vectors_{ts}.jsonl.gz"))

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

def query(vector: list, top_k: int = 5) -> list:
    vec_str = "[" + ",".join(f"{x:.8f}" for x in vector) + "]"
    conn    = _connect()
    cur     = conn.cursor()
    try:
        cur.execute(f"""
            SELECT TOP {top_k}
                v.item_id,
                COALESCE(sp.filename, s3.filename) AS filename,
                v.chunk_id, c.chunk_data,
                COSINE_SIMILARITY(v.vector_data, TO_REAL_VECTOR(?)) AS score
            FROM {_t("CATALYSTPLATFORM_BRAIN_VECTORS")} v
            JOIN {_t("CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS")} c ON v.item_id = c.item_id AND v.chunk_id = c.chunk_id
            LEFT JOIN {_t("CATALYSTPLATFORM_BRAIN_SP_FILES")} sp ON v.item_id = sp.item_id
            LEFT JOIN {_t("CATALYSTPLATFORM_BRAIN_S3_FILES")} s3 ON v.item_id = s3.item_id
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

def truncate_tables() -> dict:
    """Delete all rows from all BRAIN tables. Returns per-table row counts deleted."""
    conn = _connect()
    cur  = conn.cursor()
    tables = [
        "CATALYSTPLATFORM_BRAIN_SP_FILES",
        "CATALYSTPLATFORM_BRAIN_S3_FILES",
        "CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS",
        "CATALYSTPLATFORM_BRAIN_VECTORS",
        "CATALYSTPLATFORM_BRAIN_FILE_REGISTRY",
    ]
    result = {}
    try:
        for table in tables:
            cur.execute(f"SELECT COUNT(*) FROM {_t(table)}")
            before = cur.fetchone()[0]
            cur.execute(f"DELETE FROM {_t(table)}")
            result[table] = {"deleted": before}
        conn.commit()
    finally:
        cur.close()
    return result


def drop_tables() -> dict:
    """Drop all BRAIN tables. Tables must be recreated via /db/init before use."""
    conn = _connect()
    cur  = conn.cursor()
    tables = [
        "CATALYSTPLATFORM_BRAIN_FILES",
        "CATALYSTPLATFORM_BRAIN_SP_FILES_CHUNK",
        "CATALYSTPLATFORM_BRAIN_SP_FILES_VERSIONS",
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
    Apply schema migrations to existing tables.
    Safe to run multiple times — skips columns that already exist.
    Current migrations:
      - Add source NVARCHAR(20) DEFAULT 'sharepoint' to BRAIN_SP_FILES, TEXT_CHUNKS, VECTORS, FILES_CHUNK
      - Widen item_id from NVARCHAR(200) to NVARCHAR(500) across all tables (note: PK widening
        requires drop/recreate — skipped here; new rows with long S3 keys will fail on old deployments
        until tables are recreated. Run create_tables() on a fresh schema to get the new size.)
    """
    conn   = _connect()
    cur    = conn.cursor()
    status = {}

    # Columns to add: (table, column, definition)
    migrations = [
        ("CATALYSTPLATFORM_BRAIN_SP_FILES",       "source", "NVARCHAR(20) DEFAULT 'sharepoint'"),
        ("CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS",  "source", "NVARCHAR(20) DEFAULT 'sharepoint'"),
        ("CATALYSTPLATFORM_BRAIN_VECTORS",      "source", "NVARCHAR(20) DEFAULT 'sharepoint'"),
    ]

    try:
        # Fetch existing columns once
        cur.execute(
            "SELECT TABLE_NAME, COLUMN_NAME FROM SYS.TABLE_COLUMNS "
            "WHERE SCHEMA_NAME = CURRENT_SCHEMA"
        )
        existing_cols = {(row[0], row[1]) for row in cur.fetchall()}

        for table, column, definition in migrations:
            key = (table, column.upper())
            if key in existing_cols:
                status[f"{table}.{column}"] = "already exists"
                log.info(f"{_t(table)}.{column} — already exists")
            else:
                cur.execute(
                    f"ALTER TABLE {_t(table)} ADD ({column} {definition})"
                )
                status[f"{table}.{column}"] = "added"
                log.info(f"{_t(table)}.{column} — added")

        conn.commit()
        log.info(f"Migration complete — {status}")
        return status
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()


def get_table_stats() -> dict:
    """Return row counts for all BRAIN tables, with source breakdown where available."""
    conn = _connect()
    cur  = conn.cursor()
    try:
        # (table_name, key, has_source_col)
        tables = [
            ("CATALYSTPLATFORM_BRAIN_SP_FILES",        "BRAIN_SP_FILES",        True),
            ("CATALYSTPLATFORM_BRAIN_S3_FILES",        "BRAIN_S3_FILES",        False),
            ("CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS",     "BRAIN_TEXT_CHUNKS",     True),
            ("CATALYSTPLATFORM_BRAIN_VECTORS",         "BRAIN_VECTORS",         True),
            ("CATALYSTPLATFORM_BRAIN_FILES_CHUNK",     "BRAIN_FILES_CHUNK",     True),
            ("CATALYSTPLATFORM_BRAIN_FILE_REGISTRY",   "BRAIN_FILE_REGISTRY",   True),
            ("CATALYSTPLATFORM_BRAIN_SP_FILES_VERSIONS","BRAIN_SP_FILES_VERSIONS",False),
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
            f"SELECT item_id, last_modified FROM {_t('CATALYSTPLATFORM_BRAIN_SP_FILES')}"
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
            f"FROM {_t('CATALYSTPLATFORM_BRAIN_SP_FILES')} ORDER BY indexed_at DESC"
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



def get_db_analysis() -> dict:
    """Collect all metrics for the DB analysis dashboard."""
    conn = _connect()
    cur  = conn.cursor()
    try:
        bf  = _t("CATALYSTPLATFORM_BRAIN_SP_FILES")
        btc = _t("CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS")
        bv  = _t("CATALYSTPLATFORM_BRAIN_VECTORS")
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
        bs3 = _t("CATALYSTPLATFORM_BRAIN_S3_FILES")

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
    conn = _connect()
    cur  = conn.cursor()
    try:
        cur.execute(f"DELETE FROM {_t('CATALYSTPLATFORM_BRAIN_VECTORS')}        WHERE item_id = ?", (item_id,))
        cur.execute(f"DELETE FROM {_t('CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS')}    WHERE item_id = ?", (item_id,))
        cur.execute(f"DELETE FROM {_t('CATALYSTPLATFORM_BRAIN_SP_FILES')}          WHERE item_id = ?", (item_id,))
        conn.commit()
        log.info(f"Deleted item {item_id}")
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
    top_k: int = Field(5, ge=1, le=50, description="Number of results to return")


@router.post("/query")
def db_query(req: QueryRequest):
    from api import _get_embed_model
    try:
        vector  = _get_embed_model().encode(req.text).tolist()
        results = query(vector, top_k=req.top_k)
        return {"query": req.text, "count": len(results), "results": results}
    except Exception as e:
        log.error(f"Query failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


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




@router.delete("/item/{item_id}", dependencies=[Depends(_require_destructive)])
def db_delete(item_id: str):
    try:
        return delete_item(item_id)
    except Exception as e:
        log.error(f"DB delete failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/init")
def db_init():
    """Create all BRAIN tables if they do not already exist. Safe to run multiple times."""
    try:
        result = create_tables()
        return {"status": "ok", "tables": result}
    except Exception as e:
        log.error(f"DB init failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/migrate")
def db_migrate():
    """
    Apply schema migrations to existing tables (safe to run multiple times).
    Adds: source column to BRAIN_SP_FILES, TEXT_CHUNKS, VECTORS, FILES_CHUNK.
    """
    try:
        result = migrate_tables()
        return {"status": "ok", "migrations": result}
    except Exception as e:
        log.error(f"DB migration failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/drop", dependencies=[Depends(_require_destructive)])
def db_drop(confirm: bool = False):
    """
    Drop all 5 BRAIN tables entirely (schema + data).
    Call POST /db/init afterward to recreate them.
    Requires ?confirm=true to prevent accidental calls.
    """
    if not confirm:
        raise HTTPException(
            status_code=400,
            detail="Pass ?confirm=true to confirm dropping all BRAIN tables.",
        )
    try:
        result = drop_tables()
        dropped = [t for t, s in result.items() if s == "dropped"]
        log.warning(f"DB drop executed — {len(dropped)} tables dropped")
        return {"status": "ok", "tables": result}
    except Exception as e:
        log.error(f"DB drop failed: {type(e).__name__}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/truncate", dependencies=[Depends(_require_destructive)])
def db_truncate(confirm: bool = False):
    """
    Delete ALL rows from all BRAIN tables.
    Requires ?confirm=true to prevent accidental calls.
    """
    if not confirm:
        raise HTTPException(
            status_code=400,
            detail="Pass ?confirm=true to confirm deletion of all rows from all BRAIN tables.",
        )
    try:
        result = truncate_tables()
        total = sum(v["deleted"] for v in result.values())
        log.warning(f"DB truncate executed — {total} total rows deleted across {len(result)} tables")
        return {"status": "ok", "total_deleted": total, "tables": result}
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
                f"FROM {_t('CATALYSTPLATFORM_BRAIN_FILE_REGISTRY')} "
                f"WHERE source = ? ORDER BY insert_datetime DESC LIMIT ? OFFSET ?",
                (source, limit, offset)
            )
        else:
            cur.execute(
                f"SELECT item_id, modified_datetime, source, filename, insert_datetime "
                f"FROM {_t('CATALYSTPLATFORM_BRAIN_FILE_REGISTRY')} "
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
