# SharePoint RAG API — Technical Specification

> **Path:** `cf8/push/`  
> **Entry point:** `api.py` (FastAPI + uvicorn)  
> **Runtime:** SAP BTP Cloud Foundry, Python buildpack  
> **Live URL:** `atcp-sap-ai-delivery-workbench-sap-ai-workbench-sharepo44c7e0a9.cfapps.ap10.hana.ondemand.com`

---

## 1. Overview

A FastAPI service that ingests documents from three sources — **Microsoft SharePoint** (via MS Graph API), **SAP BTP Object Store (S3)**, and the **local `/tmp` filesystem** — extracts text, semantically chunks and embeds the content, and stores vectors in **SAP HANA Cloud** for retrieval-augmented generation (RAG) queries.

### Services Bound (CF)

| Service binding | Purpose |
|---|---|
| `btp-catalyst-platform-db` | HANA HDI container — injects `VCAP_SERVICES.hana` credentials |
| `btp-catalyst-platform-destination` | BTP Destination Service — resolves MS Graph OAuth2 credentials from destination `BTP-CP_MS_Graph` |
| `btp-catalyst-platform-objectstore` | SAP BTP Object Store (S3-compatible) — injects `VCAP_SERVICES.objectstore` credentials |

### Resource limits (manifest.yml)

| Parameter | Value |
|---|---|
| Memory | 1536M |
| Disk quota | 4096M |
| Instances | 1 |
| Health check | HTTP `GET /health` (120s invocation timeout) |
| Start timeout | 180s |

---

## 2. Architecture

```
┌─────────────────────────┐
│  Source (one of three)  │
│  1. SharePoint (Graph)  │
│  2. S3 (Object Store)   │
│  3. /tmp filesystem     │
└──────────┬──────────────┘
           │  raw bytes
           ▼
      _extract_text()          ← docx/xlsx/pdf/pptx/legacy/zip/text
           │  plain text
           ▼
     _semantic_chunks()        ← paragraph-based, max_tokens=512 (⚠ see §7)
           │  list[str]
           ▼
  _get_embed_model().encode()  ← BAAI/bge-small-en-v1.5, 384-dim, normalize_embeddings=True
           │  list[float] per chunk
           ▼
       db.ingest()             ← UPSERT BRAIN_SP/S3_FILES + TEXT_CHUNKS + VECTORS
           │
           ▼
  _upload_run_artifacts()      ← HTML report + call log → S3 runs/<YYYYMMDD>/
```

### Module responsibilities

| Module | Role |
|---|---|
| `api.py` | FastAPI app, all route handlers, extraction, chunking, embedding |
| `db.py` | HANA connection, UPSERT helpers, vector similarity query, router `/db/*` |
| `s3.py` | boto3 S3 client wrapper — list, upload, download, delete, presigned URLs |

### Tech stack

| Layer | Library / version |
|---|---|
| API framework | `fastapi>=0.110.0`, `uvicorn[standard]>=0.27.0` |
| HTTP | `requests>=2.31.0`, `urllib3>=2.0.0` |
| DOCX | `python-docx>=1.1.0` |
| XLSX | `openpyxl>=3.1.2`, `xlrd==1.2.0` |
| PDF | `pdfplumber>=0.10.0` |
| PPTX | `python-pptx>=0.6.23` |
| Legacy Word/OLE | `olefile>=0.46` |
| RTF | `striprtf>=0.0.26` |
| MSG (Outlook) | `extract-msg>=0.28.0` |
| Embeddings | `sentence-transformers>=2.7.0`, `torch>=2.1.0` (CPU-only build) — model: `BAAI/bge-small-en-v1.5` |
| HANA | `hdbcli>=2.28.17` |
| S3 | `boto3>=1.34.0` |

---

## 3. Database Schema (SAP HANA Cloud)

All tables are prefixed `CATALYSTPLATFORM_` and are created by `db.create_tables()`.  
The schema is resolved from `VCAP_SERVICES` credentials at startup; all DDL uses `_t()` to qualify names.

### CATALYSTPLATFORM_BRAIN_SP_FILES

Metadata for SharePoint files.

| Column | Type | Notes |
|---|---|---|
| `item_id` | `NVARCHAR(500)` | PK — Graph API item ID |
| `filename` | `NVARCHAR(500)` | |
| `file_type` | `NVARCHAR(200)` | |
| `extension` | `NVARCHAR(20)` | |
| `size_text` | `NVARCHAR(50)` | Human-readable, e.g. `"2.3 MB"` |
| `size_bytes` | `BIGINT` | |
| `web_url` | `NVARCHAR(1000)` | |
| `doc_path` | `NVARCHAR(1000)` | Graph download URL |
| `last_modified` | `NVARCHAR(50)` | ISO 8601 string |
| `created_date` | `NVARCHAR(50)` | |
| `created_by` | `NVARCHAR(500)` | |
| `modified_by` | `NVARCHAR(500)` | |
| `shared` | `NVARCHAR(100)` | |
| `version_count` | `INT` | |
| `source` | `NVARCHAR(20)` | DEFAULT `'sharepoint'` |
| `indexed_at` | `TIMESTAMP` | DEFAULT `CURRENT_TIMESTAMP` |
| `insert_dt` | `TIMESTAMP` | Set by ingest |

### CATALYSTPLATFORM_BRAIN_S3_FILES

Metadata for S3 objects.

| Column | Type | Notes |
|---|---|---|
| `item_id` | `NVARCHAR(500)` | PK — S3 ETag (MD5) |
| `s3_key` | `NVARCHAR(1000)` | Full object key |
| `filename` | `NVARCHAR(500)` | Basename of key |
| `extension` | `NVARCHAR(20)` | |
| `size_text` | `NVARCHAR(50)` | Human-readable |
| `size_bytes` | `BIGINT` | |
| `etag` | `NVARCHAR(200)` | MD5 of object content (S3 ETag without quotes) |
| `storage_class` | `NVARCHAR(50)` | e.g. `"STANDARD"` |
| `last_modified` | `NVARCHAR(50)` | ISO 8601 string |
| `indexed_at` | `TIMESTAMP` | DEFAULT `CURRENT_TIMESTAMP` |
| `insert_dt` | `TIMESTAMP` | Set by ingest |

> **Note:** `/tmp` files use source `"tmp"` and write to `BRAIN_SP_FILES` (not a separate table). `item_id` = `md5(abs_path)`.

### CATALYSTPLATFORM_BRAIN_TEXT_CHUNKS

Semantic text chunks.

| Column | Type | Notes |
|---|---|---|
| `item_id` | `NVARCHAR(500)` | PK (part 1) |
| `chunk_id` | `INT` | PK (part 2) — 1-based sequence |
| `chunk_data` | `NCLOB` | Raw chunk text |
| `filename` | `NVARCHAR(500)` | |
| `source` | `NVARCHAR(20)` | `'sharepoint'`, `'s3'`, or `'tmp'` |
| `insert_dt` | `TIMESTAMP` | |

### CATALYSTPLATFORM_BRAIN_VECTORS

384-dimensional sentence embeddings.

| Column | Type | Notes |
|---|---|---|
| `item_id` | `NVARCHAR(500)` | PK (part 1) |
| `chunk_id` | `INT` | PK (part 2) |
| `vector_data` | `REAL_VECTOR(384)` | Stored via `TO_REAL_VECTOR(?)` |
| `filename` | `NVARCHAR(500)` | |
| `source` | `NVARCHAR(20)` | |
| `insert_dt` | `TIMESTAMP` | |

### CATALYSTPLATFORM_BRAIN_FILES_VERSIONS

Version history for SharePoint files, populated by `/files/versions` and `/pipeline`.

| Column | Type | Notes |
|---|---|---|
| `item_id` | `NVARCHAR(500)` | PK (part 1) |
| `version_id` | `NVARCHAR(50)` | PK (part 2) |
| `file_size_text` | `NVARCHAR(50)` | Human-readable size |
| `size_bytes` | `BIGINT` | |
| `last_modified_dt` | `NVARCHAR(50)` | ISO 8601 string |
| `lastmodified_by` | `NVARCHAR(500)` | |
| `download_url` | `NCLOB` | Pre-authenticated version download URL |
| `version_count` | `INT` | Total versions for this file |
| `filename` | `NVARCHAR(500)` | |

### CATALYSTPLATFORM_BRAIN_FILES_CHUNK

Per-file chunk count summary, populated after every successful ingest by `db.ingest_chunk_summary()`.

| Column | Type | Notes |
|---|---|---|
| `item_id` | `NVARCHAR(500)` | PK |
| `filename` | `NVARCHAR(500)` | |
| `chunk_count` | `INT` | Total chunks produced for this file |
| `source` | `NVARCHAR(20)` | `'sharepoint'`, `'s3'`, or `'tmp'` |
| `insert_dt` | `TIMESTAMP` | |

### CATALYSTPLATFORM_BRAIN_FILE_REGISTRY

Change-detection registry — one row per file, updated after every successful ingest.

| Column | Type | Notes |
|---|---|---|
| `item_id` | `NVARCHAR(500)` | PK |
| `modified_datetime` | `NVARCHAR(50)` | Last-modified timestamp that was ingested |
| `source` | `NVARCHAR(20)` | |
| `filename` | `NVARCHAR(500)` | |
| `insert_datetime` | `TIMESTAMP` | DEFAULT `CURRENT_TIMESTAMP` |

> **Note:** `BRAIN_SP_FILES.last_modified` is used as the actual change-detection key in `_filter_stale_items()` — `db.get_modified_map()` reads from `BRAIN_SP_FILES`. `BRAIN_FILE_REGISTRY` is written after every successful ingest (best-effort, non-fatal) but is **not** the source the filter reads. `/tmp` files also write to `BRAIN_SP_FILES` (source=`'tmp'`), not to a separate table.

---

## 4. Authentication

### SharePoint / MS Graph

| Environment | Token source |
|---|---|
| SAP BTP CF | BTP Destination Service → destination `BTP-CP_MS_Graph` → Azure AD OAuth2 client credentials |
| Local dev | `TENANT_ID` / `CLIENT_ID` / `CLIENT_SECRET` env vars → `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token` |

Tokens are cached in `_token_cache` and auto-refreshed 60 s before expiry. The dependency `require_connection` blocks all SharePoint endpoints until `drive_id` is resolved at startup.

---

## 5. API Endpoints

Interactive docs: `GET /docs` (Swagger UI), `GET /redoc`.

### Tag: system

| Method | Path | Description |
|---|---|---|
| `GET` | `/health` | Application health check. Returns `{"status": "ok", "timestamp": "..."}` |
| `GET` | `/debug/destination` | Diagnose BTP Destination Service connectivity. Requires `ENABLE_DEBUG=true`. |
| `GET` | `/debug/db` | Diagnose HANA connectivity. Requires `ENABLE_DEBUG=true`. |
| `GET` | `/debug/s3` | Diagnose S3 / Object Store connectivity. Requires `ENABLE_DEBUG=true`. |
| `GET` | `/debug/tmp` | Browse the `/tmp` filesystem one level at a time. Query param: `path`. |
| `GET` | `/debug/tmp/preview` | Preview a text file from `/tmp`. Query param: `path`. |

### Tag: files (SharePoint)

All `/files` endpoints require SharePoint connection (`require_connection`).

| Method | Path | Description |
|---|---|---|
| `GET` | `/files` | List all root files + sync metadata to HANA. Response: `FilesResponse`. |
| `GET` | `/files/versions` | List root files with full version history. |
| `GET` | `/files/all` | List files from a specific folder path (or root). Query: `folder_path`. |
| `GET` | `/files/search` | Search SharePoint files by filename. Query: `name` (partial match via Graph search). |
| `POST` | `/download` | Download a SharePoint file and stream it back. Body: `{"item": {...}}`. |
| `GET` | `/rebuild` | Rebuild a file from stored chunk records (reads `files_list_chunk.json`). |
| `POST` | `/rebuild` | Rebuild a file from a supplied chunk list. Body: `{"chunks": [...]}`. |

### Tag: rag (Vectorization)

#### SharePoint sources

| Method | Path | Description |
|---|---|---|
| `GET` | `/vectorize` | Full pipeline — all root files, delta-based change detection. |
| `GET` | `/vectorize/folder/{folder_path}` | Same pipeline scoped to a SharePoint subfolder. Query: `delay`, `use_delta`. |
| `POST` | `/vectorize` | Vectorize an explicit item list (or auto-fetch root if `items` omitted). Body: `VectorizeRequest`. |
| `POST` | `/vectorize/item` | Vectorize a single SharePoint item by Graph item ID. Body: `{"item_id": "..."}`. |
| `POST` | `/pipeline` | Full pipeline with version history guaranteed before vectorize. Body: `PipelineRequest`. |

#### S3 sources

| Method | Path | Description |
|---|---|---|
| `POST` | `/vectorize/s3` | Vectorize all files under an S3 prefix. Body: `S3VectorizeRequest`. |
| `GET` | `/vectorize/s3/folder/{folder_path}` | Same as `POST /vectorize/s3` but prefix in URL. Query: `force`, `extensions`. |
| `GET` | `/vectorize/s3/folder/{folder_path}/download` | Download all files under a prefix as a ZIP. `skip_ingest=True`. |
| `POST` | `/vectorize/s3/item` | Vectorize a single S3 object by key. Body: `{"key": "...", "force": false}`. |

#### /tmp (local filesystem) sources

| Method | Path | Description |
|---|---|---|
| `POST` | `/vectorize/tmp` | Vectorize all supported files under a local folder. Body: `TmpVectorizeRequest`. |
| `GET` | `/vectorize/tmp/{folder_path}` | Same but folder appended to `/tmp/` from URL path. Query: `force`, `extensions`. |
| `POST` | `/vectorize/tmp/item` | Vectorize a single local file by absolute path. Body: `TmpVectorizeItemRequest`. |

### Tag: s3

| Method | Path | Description |
|---|---|---|
| `GET` | `/s3/objects` | List objects in the bucket. Query: `prefix`, `delimiter`. |
| `GET` | `/s3/object/{s3_key}/metadata` | Get metadata for an S3 object. |
| `GET` | `/s3/object/{s3_key}` | Stream-download an S3 object through the API. |
| `GET` | `/s3/presigned/{s3_key}` | Generate a presigned GET URL. Query: `expiry`. |
| `POST` | `/s3/presigned-upload` | Generate a presigned POST policy for direct upload. Body: `S3PresignedUploadRequest`. |
| `DELETE` | `/s3/object/{s3_key}` | Delete a single S3 object. Requires `ENABLE_DESTRUCTIVE=true`. |
| `DELETE` | `/s3/objects` | Bulk delete up to 1000 objects. Body: `{"keys": [...]}`. |
| `POST` | `/s3/upload-storage-file` | Upload a local output file to S3. Body: `S3UploadStorageRequest`. |
| `POST` | `/s3/upload-storage-folder/{day}` | Bulk upload an entire date folder to S3. |

### Tag: storage

| Method | Path | Description |
|---|---|---|
| `GET` | `/storage` | List all local output files by date folder. |
| `GET` | `/storage/{day}` | Browse a date folder as HTML. |
| `GET` | `/storage/{day}/{filename}` | View or download a specific output file. |
| `DELETE` | `/storage/{day}/{filename}` | Delete a single output file. Requires `ENABLE_DESTRUCTIVE=true`. |
| `DELETE` | `/storage/{day}` | Delete an entire date folder. Requires `ENABLE_DESTRUCTIVE=true`. |

### Tag: analysis

| Method | Path | Description |
|---|---|---|
| `GET` | `/analysis` | List all run reports and logs (by date folder). |
| `GET` | `/analysis/latest` | View the most recent run report as HTML. |
| `GET` | `/analysis/db` | Live DB metrics dashboard — row counts and stats for all HANA BRAIN tables, rendered as HTML. |
| `GET` | `/analysis/{day}/{filename}` | View a specific report or log. |

### Tag: database (router prefix `/db`)

| Method | Path | Description |
|---|---|---|
| `POST` | `/db/query` | **Layer-filtered** vector similarity search. Body: `{"text": "...", "top_k": 5, "layer": ..., "deliverable_code": ..., "min_score": 0.25}`. Returns top-k chunks with cosine scores and their tags. See [Filtered retrieval](#filtered-retrieval). |
| `POST` | `/db/chunks` | **Deterministic** tag-only fetch — no query text, no ranking. Returns every matching chunk ordered by `filename, chunk_id`. At least one filter is required (a filterless call returns 400 rather than the whole KB). Used to pull a deliverable's template whole and in order. |
| `GET` | `/db/tables` | Which BRAIN table set is active (`read_set`, `write_sets`) plus per-table presence and row counts for **both** the legacy `BRAIN_*` and HDI `TB_BRAIN_*` sets, with a `matched` flag per table. Read-only. |
| `GET` | `/db/vectors/{item_id}` | Return decoded `REAL_VECTOR(384)` rows for a given `item_id`. Response includes `chunk_id` and `vector` float array (384 elements). |
| `POST` | `/db/sql` | Run a validated **SELECT-only** SQL query against HANA. Body: `SqlRequest`. Guards: first keyword must be `SELECT`; blocks DML/DDL, `--`/`/* */` comments, semicolons, and dangerous keywords (`DROP`, `DELETE`, `INSERT`, `UPDATE`, `EXEC`, etc.). |
| `GET` | `/db/files` | List all rows in the active SP files table. |
| `GET` | `/db/chunks/{item_id}` | Get all chunks for a file. |
| `GET` | `/db/versions` | List all version rows. |
| `GET` | `/db/versions/{item_id}` | Get version history for a specific file. |
| `GET` | `/db/chunk-summary` | List all raw-chunk rows (one row per file chunk). |
| `GET` | `/db/chunk-summary/{item_id}` | Get chunk count summary for a specific file. |
| `DELETE` | `/db/item/{item_id}` | Delete a file + its chunks + vectors, from **every** configured write set. Requires `ENABLE_DESTRUCTIVE=true`. |
| `POST` | `/db/init` | Create the **legacy** `BRAIN_*` tables (idempotent). Returns **409** when `BRAIN_READ_SET=hdi` — the `TB_BRAIN_*` tables are owned by the HDI `db` module and must be created by deploying it. |
| `POST` | `/db/migrate` | Additive column migrations on the **legacy** tables (idempotent): adds `source` plus the `layer` / `layer_category` / `deliverable_code` / `project` retrieval tags to SP_FILES, S3_FILES, TEXT_CHUNKS and VECTORS. Returns **409** when `BRAIN_READ_SET=hdi`. |
| `POST` | `/db/drop` | Drop the orphaned pre-HDI tables. Requires `ENABLE_DESTRUCTIVE=true` + `?confirm=true`. Returns **409** when `BRAIN_READ_SET=hdi`. |
| `POST` | `/db/truncate` | Delete all rows. Requires `ENABLE_DESTRUCTIVE=true` + `?confirm=true`. Clears every configured write set; pass `?table_set=hdi` to reset only the HDI duplicates between test runs. |
| `GET` | `/db/file-registry` | List file-registry rows. Query: `source`, `limit`, `offset`. |
| `POST` | `/db/backup` | Dump all BRAIN tables to JSONL and upload to S3. Body: `{"s3_dir": "...", "include_vectors": true}`. |
| `POST` | `/db/backfill` | Copy rows that exist in the **legacy** `BRAIN_*` tables but are missing from the **HDI** `TB_BRAIN_*` tables (INSERT … WHERE NOT EXISTS on PK — idempotent). Body: `{}` for all tables or `{"table": "vectors"}` for one. Returns per-table inserted count; skips tables that don't exist yet. Requires the HDI db module to be deployed first. |

---

### Table sets — the HDI migration

The BRAIN tables exist twice while ownership moves from Python to the HDI `db`
module in BTP-Catalyst-Platform:

| Set | Tables | Owner |
|---|---|---|
| `legacy` | `CATALYSTPLATFORM_BRAIN_*` | created by `db.py` (`/db/init`) |
| `hdi` | `CATALYSTPLATFORM_TB_BRAIN_*` | `db/src/*.hdbmigrationtable` in BTP-Catalyst-Platform |

Two env vars select which set is used. Both default to `legacy`, so a plain
`cf push` behaves exactly as before.

| Variable | Values | Default | Effect |
|---|---|---|---|
| `BRAIN_READ_SET` | `legacy` \| `hdi` | `legacy` | Which set every SELECT reads from. |
| `BRAIN_WRITE_SET` | `legacy` \| `hdi` \| `both` | `legacy` | Which set(s) receive writes. |

**`BRAIN_WRITE_SET=both` is the migration-testing mode:** every ingest lands in
both sets, so the HDI tables fill with real data from ordinary API traffic while
the legacy tables stay authoritative. Writes to the *read* set are primary and a
failure there raises as usual; writes to the other set are best-effort — failures
are logged and returned in the ingest result under `errors`, never breaking the
live path. `GET /db/tables` then shows both sets side by side with a `matched`
flag so you can confirm they agree.

Python issues **no DDL** against the HDI set. `/db/init`, `/db/migrate` and
`/db/drop` return 409 when `BRAIN_READ_SET=hdi`.

---

### Filtered retrieval

`POST /db/query` accepts these filters alongside `text` and `top_k`. Each takes a
single value or a list (list → SQL `IN`). Filters are applied **before** the
cosine ranking, so `top_k` counts rows that already matched.

| Filter | Example |
|---|---|
| `layer` | `"Layer 1 - SAP Guidelines"` or `["Layer 1 - SAP Guidelines", "Layer 2 - Accenture Guidelines"]` |
| `layer_category` | `"Templates"` |
| `deliverable_code` | `"B10"` — deliverable rowId, from `deliverables.json` |
| `project` | `"ACME-2026"`; `null` for global reference material |
| `source` | `"sharepoint"` \| `"s3"` |
| `min_score` | `0.25` — cosine floor; weaker matches are dropped |

An unknown filter column returns **400** rather than being ignored — a typo that
silently matched everything would hand back the whole knowledge base, which is
the failure this endpoint exists to prevent.

Tags are derived from the BTP-Brain folder path at vectorize time
(`_parse_s3_tags` / `_flatten_item` in `api.py`):
`knowledge-base/<layer>/<layer_category>/<deliverable>/<file>`. The deliverable
folder is matched case-insensitively against the short label, rowId or full name
in `deliverables.json`. Explicit overrides can be passed on the vectorize
requests (`layer`, `layer_category`, `deliverable_code`, `project`) when the S3
layout does not follow that convention.

**Consumer:** `srv/lib/brain-retrieval.js` in BTP-Catalyst-Platform calls
`/db/chunks` for the deliverable's template (deterministic, whole, in order) and
`/db/query` for its guidelines (semantic, multi-query, relevance-ranked), then
assembles both into the reference block injected into generation and findings
prompts. Node never embeds anything — one model on both sides keeps the
similarity scores comparable.

---

## 6. Key Request/Response Models

### VectorizeResponse
```json
{
  "count_chunks":      42,
  "count_files":       3,
  "skipped_unchanged": 7,
  "saved_to":          ["/tmp/20260907/sp_chunks_...", "..."],
  "db_ingest":         {"per_file_ingest": [...]}
}
```

### S3VectorizeRequest
```json
{
  "prefix":     "BTP-Catalyst/knowledge-base/",
  "force":      false,
  "extensions": ["pdf", "docx"]
}
```

### TmpVectorizeRequest
```json
{
  "folder":     "/tmp/uploads",
  "force":      false,
  "extensions": ["pdf", "docx"]
}
```

### TmpVectorizeItemRequest
```json
{
  "path":  "/tmp/uploads/report.pdf",
  "force": false
}
```

### db.QueryRequest → db.query() response
```json
{
  "query": "what is the project timeline",
  "count": 5,
  "results": [
    {
      "item_id":    "abc123",
      "filename":   "ProjectPlan.docx",
      "chunk_id":   4,
      "chunk_data": "...",
      "score":      0.8921
    }
  ]
}
```

### GET /db/vectors/{item_id} response
```json
{
  "item_id": "01HOEYRQAGJHFYDPTVX5CZOAOFUZNH6ZGI",
  "count": 10,
  "vectors": [
    { "chunk_id": 1, "vector": [-0.0103, -0.0217, 0.0356, "...384 floats total"] }
  ]
}
```

> **Note on binary format:** The `REAL_VECTOR(384)` column is returned by hdbcli as raw binary bytes (2-byte `num_dims` LE + 2-byte padding + 384 × float32 LE = 1540 bytes). The endpoint decodes this to a plain float list before serialising to JSON.

### SqlRequest (POST /db/sql)
```json
{
  "sql":    "SELECT a.col1, b.col2 FROM \"SCHEMA\".\"TABLE_A\" a LEFT JOIN \"SCHEMA\".\"TABLE_B\" b ON a.col1 = b.col1 WHERE a.col1 = ? AND b.col2 = ?",
  "params": ["value1", "value2"],
  "limit":  100
}
```

### SqlRequest → response
```json
{
  "count":   2,
  "columns": ["COL1", "COL2"],
  "rows": [
    { "COL1": "value1", "COL2": "value2" },
    { "COL1": "value1", "COL2": "value3" }
  ]
}
```

> **SQL injection guards (enforced server-side):**
> - First keyword must be `SELECT` — any other statement is rejected with HTTP 400
> - Blocked keywords anywhere in the query: `INSERT`, `UPDATE`, `DELETE`, `DROP`, `CREATE`, `ALTER`, `TRUNCATE`, `EXEC`, `EXECUTE`, `GRANT`, `REVOKE`, `MERGE`, `INTO`, `OUTFILE`, `LOAD`
> - `--` and `/* */` comment sequences → HTTP 400
> - Semicolons (statement stacking) → HTTP 400
> - Values must be passed via `params` list using `?` placeholders — never interpolated into the SQL string

---

## 7. Vectorization Pipeline — Current Implementation

### Step 1: Input (source-specific)

| Source | Function | Input |
|---|---|---|
| SharePoint | `_download_and_vectorize(items)` | List of Graph API item dicts with `BaseURL` (download URL) |
| S3 | `_download_and_vectorize_s3(objects)` | List of S3 object dicts from `s3.list_objects()` |
| /tmp | `_download_and_vectorize_tmp(paths)` | List of absolute local file paths |

SharePoint download uses `requests.get(url, stream=True)` with 3 retries, exponential backoff (5/15/30s), 429 Retry-After handling.

### Step 2: Extraction — `_extract_text(content: bytes, filename: str) -> str`

| Format | Library | Method |
|---|---|---|
| `.docx`, `.dotx` | `python-docx` | `doc.paragraphs` joined |
| `.doc` | `olefile` | `WordDocument` stream, UTF-16-LE decoded |
| `.xlsx`, `.xlsm` | `openpyxl` | `iter_rows(values_only=True)`, `\|`-delimited |
| `.xls` | `xlrd` | `cell_value()`, `\|`-delimited |
| `.pdf` | `pdfplumber` | `page.extract_text()` per page |
| `.pptx`, `.pptm` | `python-pptx` | `shape.text` per slide shape |
| `.rtf` | `striprtf` | `rtf_to_text()` |
| `.msg` | `extract-msg` | `msg.body` |
| `.csv` | stdlib | `csv.DictReader`, row dicts joined |
| `.json` | stdlib | `json.dumps(indent=2)` |
| `.xml`, `.html`, `.htm` | stdlib regex | Tag strip |
| `.zip` | stdlib `zipfile` | Recursive extraction, each inner file re-processed |
| Text types (`.txt`, `.md`, `.py`, `.js`, `.ts`, `.yaml`, `.yml`, `.ini`, `.toml`, `.log`, `.cfg`, `.sql`, `.java`, `.cs`, `.abap`) | stdlib | `decode('utf-8', errors='replace')` |

### Step 3: Chunking — `_semantic_chunks(text: str, max_tokens: int = 512) -> list[str]`

```python
max_chars  = max_tokens * 4          # token ≈ 4 chars (character approximation, NOT tokenizer)
paragraphs = [p.strip() for p in text.split("\n") if p.strip()]
# Accumulate paragraphs until max_chars exceeded, then flush to a new chunk
```

- **Chunk size:** `max_tokens=512` hardcoded at every call site → `max_chars=2048`
- **Token counting:** character approximation (`chars / 4`), not actual wordpiece tokenizer
- **Overlap:** **none** — hard cut at paragraph boundary
- **Metadata:** **none attached** — each chunk is a plain string; `item_id`, `filename`, `chunk_id` are added by the caller

### Step 4: Embedding — `_get_embed_model()`

```python
_embed_model = SentenceTransformer("BAAI/bge-small-en-v1.5")
# Called per-chunk:
embedding = _get_embed_model().encode(chunk_text, normalize_embeddings=True).tolist()
```

- **Model:** `BAAI/bge-small-en-v1.5` — 384-dim, `max_seq_length=512` word-piece tokens, `normalize_embeddings=True` (unit-norm vectors, required for correct cosine similarity)
- **Batching:** **none** — `encode()` called one chunk at a time in a loop
- **Vector stamping:** **none** — `embed_model`, `embed_dim`, `corpus_version` not stored in `BRAIN_VECTORS`
- **Model cache:** stored in `SENTENCE_TRANSFORMERS_HOME=/tmp/models` (avoids re-download on restart)

### Step 5: HANA Upsert — `db.ingest(chunks_data, vectors_data, items, source)`

Three UPSERTs per file (wrapped in a single transaction):
1. `BRAIN_SP_FILES` (SharePoint/tmp source) or `BRAIN_S3_FILES` (S3 source)
2. `BRAIN_TEXT_CHUNKS` — one row per chunk
3. `BRAIN_VECTORS` — one row per vector, stored as `TO_REAL_VECTOR('[x,x,x,...]')`

Followed by a separate `BRAIN_FILE_REGISTRY` upsert (best-effort, non-fatal on failure).

### Step 6: Post-run artifacts

After every batch:

| Artifact | Uploaded to S3? |
|---|---|
| HTML run report (`_generate_analysis_html`) | Yes → `runs/<YYYYMMDD>/<endpoint>_analysis_<ts>.html` |
| Per-call log file | Yes → `runs/<YYYYMMDD>/<endpoint>_<ts>.log` |

---

## 8. Output Files (written to `DOWNLOAD_PATH/<YYYYMMDD>/`)

### SharePoint pipeline

| Filename pattern | Format | Contents |
|---|---|---|
| `sp_chunks_<ts>.jsonl` | JSONL | `{item_id, filename, chunk_id, chunk_data}` per chunk |
| `sp_vectors_<ts>.jsonl` | JSONL | `{item_id, filename, chunk_id, vector_data: [float×384]}` per chunk |
| `sp_files_chunk_<ts>.jsonl` | JSONL | `{item_id, filename, chunk_count}` per file |
| `sp_root_data_<ts>.json` | JSON array | Flat file metadata (pipeline and folder endpoints) |
| `sp_files_versions_<ts>.json` | JSON array | Version history records |
| `sp_analysis_<ts>.html` | HTML | Run summary report |
| `<endpoint>_<ts>.log` | Text | Per-call structured log |

### S3 pipeline

| Filename pattern | Format | Contents |
|---|---|---|
| `s3_chunks_<ts>.jsonl` | JSONL | `{item_id, filename, chunk_id, chunk_data}` |
| `s3_vectors_<ts>.jsonl` | JSONL | `{item_id, filename, chunk_id, vector_data}` |
| `s3_files_chunks_<ts>.json` | JSON | Raw download chunks (base64-encoded, 8192-byte blocks) |
| `s3_root_data_<ts>.json` | JSON array | `{item_id, filename, chunk_count, key, size, last_modified, etag, storage_class}` |

### /tmp pipeline

| Filename pattern | Format | Contents |
|---|---|---|
| `tmp_chunks_<ts>.jsonl` | JSONL | `{item_id, filename, chunk_id, chunk_data}` |
| `tmp_vectors_<ts>.jsonl` | JSONL | `{item_id, filename, chunk_id, vector_data}` |
| `tmp_files_chunks_<ts>.json` | JSON | Raw file chunks (base64-encoded, 8192-byte blocks) |
| `tmp_root_data_<ts>.json` | JSON array | `{item_id, filename, chunk_count, path, size, last_modified}` |

---

## 9. Change Detection

`_filter_stale_items(items) -> (to_process, skipped_count)`

1. Calls `db.get_modified_map()` → `{item_id: last_modified}` from `BRAIN_SP_FILES`
2. Compares `item["last_modified_date"]` (truncated to seconds) against the DB value
3. Skips the file if timestamps match; otherwise queues for processing
4. Falls back to processing all items if the DB call fails

**item_id derivation by source:**

| Source | item_id |
|---|---|
| SharePoint | Graph API opaque item ID (e.g. `01ABCDEF...`) |
| S3 | ETag (MD5 of object content, without quotes) |
| /tmp | `md5(abs_path)` — stable across runs for the same path |

---

## 10. Environment Variables

| Variable | Default | Description |
|---|---|---|
| `TENANT_ID` | — | Azure AD tenant ID (local dev only) |
| `CLIENT_ID` | — | Azure AD app client ID (local dev only) |
| `CLIENT_SECRET` | — | Azure AD app client secret (local dev only) |
| `HOSTNAME` | `z0y8z.sharepoint.com` | SharePoint hostname |
| `SITE_PATH` | `/sites/BTP-CatalystPlatform` | SharePoint site path |
| `DOWNLOAD_PATH` | `/tmp` | Root directory for all output files |
| `SENTENCE_TRANSFORMERS_HOME` | `/tmp/models` | Model cache directory (avoids re-download) |
| `HF_HUB_DISABLE_SYMLINKS_WARNING` | `1` | Suppress HuggingFace symlink warning on CF |
| `HF_HUB_DISABLE_IMPLICIT_TOKEN` | `1` | Suppress HuggingFace token warning |
| `TOKENIZERS_PARALLELISM` | `false` | Prevent tokenizer fork warnings |
| `ENABLE_DESTRUCTIVE` | `false` | Must be `true` to enable DELETE endpoints (`/storage`, `/s3`, `/db/item`) and destructive DB operations (`/db/drop`, `/db/truncate`). Default `false` on every `cf push`. |
| `BRAIN_READ_SET` | `legacy` | Which BRAIN table set SELECTs read from: `legacy` or `hdi`. See [Table sets](#table-sets--the-hdi-migration). |
| `BRAIN_WRITE_SET` | `legacy` | Which set(s) receive writes: `legacy`, `hdi`, or `both` (dual-write, for migration testing). |
| `ENABLE_DEBUG` | `false` | Must be `true` to enable `/debug/*` endpoints |
| `PORT` | `8000` | Port uvicorn binds to (set by CF automatically) |
| `VCAP_SERVICES` | — | Injected by CF — contains all bound service credentials |
| `HANA_HOST` | — | Local dev fallback (if no VCAP_SERVICES and no hana_creds.json) |
| `HANA_PORT` | `443` | |
| `HANA_USER` | — | |
| `HANA_PASSWORD` | — | |
| `HANA_SCHEMA` | — | |
| `AWS_ACCESS_KEY_ID` | — | Local dev S3 fallback |
| `AWS_SECRET_ACCESS_KEY` | — | |
| `S3_BUCKET` | — | |
| `AWS_REGION` | `us-east-1` | |
| `S3_ENDPOINT_URL` | — | For non-AWS S3-compatible endpoints |

---

## 11. Known Gaps vs Spec (Backlog)

The following issues were identified against the target vectorization spec:

### ~~Gap 1 — Chunk size exceeds model window~~ ✅ RESOLVED
- **Was:** `all-MiniLM-L6-v2` had `max_seq_length=256`; `max_tokens=512` silently truncated every chunk beyond 256 tokens.
- **Fix applied (2026-09-19):** Model replaced with `BAAI/bge-small-en-v1.5` (`max_seq_length=512`). `max_tokens=512` now aligns exactly with the model window. Validated: all 1,174 files re-indexed, 13,863 chunks produced, all within limit.

### Gap 2 — Token counting is character-approximate (MEDIUM)
- **Current:** `max_chars = max_tokens * 4` (assumes 1 token ≈ 4 chars)
- **Problem:** Word-piece tokenization is non-linear. CJK characters, punctuation, subwords, and long technical terms all tokenize differently.
- **Fix:** Use `model.tokenizer(text, return_length=True)["length"]` for accurate counting.

### Gap 3 — No overlap between chunks (MEDIUM)
- **Current:** Hard cut at paragraph boundary, zero overlap.
- **Problem:** Meaning severed at boundaries — a sentence spanning two chunks loses context in both.
- **Fix:** Add ~10–15% token overlap (~25–38 tokens for a 256-token window). Carry the last N tokens of each chunk forward as the start of the next.

### Gap 4 — No metadata attached to chunks (MEDIUM)
- **Current:** `chunk_data` is raw text only. `BRAIN_TEXT_CHUNKS` has no `section`, `page`, `deliverable_type`, `source_path`, or `version` columns.
- **Problem:** Retrieval results have no provenance beyond `filename`.
- **Fix:** Extend `BRAIN_TEXT_CHUNKS` schema and populate `section/heading`, `page` (available from pdfplumber), `deliverable_type`, `source_path`, `version`.

### Gap 5 — Embedding is not batched (LOW/MEDIUM)
- **Current:** `model.encode(chunk_text)` called once per chunk in a Python loop.
- **Problem:** Batched inference is significantly faster (especially on CPU). SentenceTransformers `encode(list[str])` handles batching internally.
- **Fix:** Collect all chunks for a file into a list, call `model.encode(chunks_list, batch_size=32)`, then zip results.

### Gap 6 — No model/version stamping on vectors (LOW)
- **Current:** `BRAIN_VECTORS` stores only `item_id`, `chunk_id`, `vector_data`, `filename`, `source`.
- **Problem:** If the embed model changes, old and new vectors are mixed in the same index without any way to distinguish them. COSINE_SIMILARITY queries will return wrong results silently.
- **Fix:** Add `embed_model NVARCHAR(100)`, `embed_dim INT`, `corpus_version NVARCHAR(50)` columns to `BRAIN_VECTORS`. Stamp on every upsert.

---

## 12. Local Development

### HANA credentials
Create `push/hana_creds.json`:
```json
{
  "host":     "your-instance.hanacloud.ondemand.com",
  "port":     443,
  "user":     "DBADMIN",
  "password": "...",
  "schema":   "YOURSCHEMA"
}
```

### S3 credentials
Create `push/s3_creds.json`:
```json
{
  "access_key_id":     "...",
  "secret_access_key": "...",
  "bucket":            "my-bucket",
  "region":            "us-east-1",
  "endpoint_url":      ""
}
```

### SharePoint credentials (local only)
Set env vars: `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`

### Run
```bash
cd push
python -m venv .venv && .venv/Scripts/activate
pip install -r requirements.txt
python api.py          # starts on http://localhost:8000
```

---

## 13. Build & Deploy (SAP BTP MTA)

From `cf8/` (parent of `push/`):

```bash
mbt build                                                           # produces mta_archives/sharepoint-rag-mta_1.0.0.mtar
cf deploy mta_archives/sharepoint-rag-mta_1.0.0.mtar               # deploy to CF
```

Or standalone:
```bash
cd push
cf push                                                             # uses manifest.yml
```
