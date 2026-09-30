"""
MCP (Model Context Protocol) server — SSE transport for Claude Desktop.

Uses FastMCP (mcp v2.x API). Tools exposed:
  search_documents      — semantic vector search across HANA
  list_indexed_files    — all files in BRAIN_SP_FILES
  get_db_stats          — HANA health / coverage summary
  list_s3_objects       — S3 / Object Store bucket listing
  get_s3_presigned_url  — generate a 1-hour download URL for an S3 object

Mounted on the existing FastAPI app via app.mount("/mcp", ...).
Protected by the global require_auth dependency (X-API-Key header).
"""

import json
import logging

import db
import s3
from fastapi import APIRouter, Request
from mcp.server.fastmcp import FastMCP

log = logging.getLogger("mcp_server")

_mcp = FastMCP("btp-catalyst")


# ── Tools ─────────────────────────────────────────────────────────────────────

@_mcp.tool()
async def search_documents(query: str, top_k: int = 5) -> str:
    """
    Semantic search across all indexed SharePoint and S3 documents stored in
    HANA. Returns the most relevant text chunks with source file info and
    similarity scores.
    """
    from api import _get_embed_model  # deferred — avoids circular import at load time
    model  = _get_embed_model()
    vector = model.encode(query).tolist()
    results = db.query(vector, top_k)
    return json.dumps(results, indent=2, ensure_ascii=False)


@_mcp.tool()
async def list_indexed_files() -> str:
    """List all files currently indexed in the HANA database (SharePoint and S3 sources)."""
    files = db.get_files()
    return json.dumps(files, indent=2, ensure_ascii=False)


@_mcp.tool()
async def get_db_stats() -> str:
    """
    Get database health and statistics: row counts per table, vector coverage %,
    missing summaries, top files by chunk count, and recent indexing activity.
    """
    stats = db.get_db_analysis()
    total_chunks  = stats.get("total_chunks") or 0
    total_vectors = stats.get("total_vectors") or 0
    summary = {
        "total_files":            stats.get("total_files"),
        "total_chunks":           total_chunks,
        "total_vectors":          total_vectors,
        "vector_coverage_pct":    round(total_vectors / total_chunks * 100, 1) if total_chunks else 0,
        "files_missing_summary":  stats.get("files_missing_chunk_summary"),
        "chunks_without_vectors": stats.get("chunks_without_vectors"),
        "by_extension":           stats.get("by_extension"),
        "top_files_by_chunks":    stats.get("top_files_by_chunks"),
        "recent_files":           stats.get("recent_files"),
        "table_stats":            stats.get("table_stats"),
    }
    return json.dumps(summary, indent=2, ensure_ascii=False)


@_mcp.tool()
async def list_s3_objects(prefix: str = "", max_keys: int = 100) -> str:
    """List objects in the SAP BTP Object Store (S3) bucket."""
    objects = s3.list_objects(prefix=prefix, max_keys=max_keys)
    return json.dumps(objects, indent=2, ensure_ascii=False)


@_mcp.tool()
async def get_s3_presigned_url(s3_key: str) -> str:
    """Generate a presigned download URL for an S3 object. Expires after 1 hour."""
    url = s3.get_presigned_url(s3_key, expiry=3600)
    return json.dumps({"url": url, "expires_in": "1 hour"})


# ── ASGI app for mounting on FastAPI ──────────────────────────────────────────

def get_asgi_app():
    """Return the ASGI app to mount on the parent FastAPI instance."""
    try:
        return _mcp.streamable_http_app()
    except AttributeError:
        return _mcp.sse_app()
