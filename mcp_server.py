"""
MCP (Model Context Protocol) server — SSE transport for Claude Desktop.

Exposes BTP Catalyst RAG capabilities as Claude tools:
  search_documents      — semantic vector search across HANA
  list_indexed_files    — all files in the BRAIN_SP_FILES table
  get_db_stats          — HANA health / coverage summary
  list_s3_objects       — S3 / Object Store bucket listing
  get_s3_presigned_url  — generate a 1-hour download URL for an S3 object

Routes (hidden from /docs, protected by global require_auth):
  GET  /mcp/sse        — Claude Desktop SSE stream
  POST /mcp/messages/  — tool call dispatcher
"""

import json
import logging

import db
import s3
from fastapi import APIRouter, Request
from mcp.server import Server
from mcp.server.sse import SseServerTransport
from mcp.types import TextContent, Tool

log = logging.getLogger("mcp_server")

_mcp = Server("btp-catalyst")
_sse = SseServerTransport("/mcp/messages/")


# ── Tool registry ─────────────────────────────────────────────────────────────

@_mcp.list_tools()
async def _list_tools():
    return [
        Tool(
            name="search_documents",
            description=(
                "Semantic search across all indexed SharePoint and S3 documents stored in HANA. "
                "Returns the most relevant text chunks with source file info and similarity scores."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {"type": "string",  "description": "Natural language search query"},
                    "top_k": {"type": "integer", "description": "Number of results to return (default: 5)", "default": 5},
                },
                "required": ["query"],
            },
        ),
        Tool(
            name="list_indexed_files",
            description="List all files currently indexed in the HANA database (SharePoint and S3 sources).",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
        Tool(
            name="get_db_stats",
            description=(
                "Get database health and statistics: row counts per table, vector coverage %, "
                "missing summaries, top files by chunk count, and recent indexing activity."
            ),
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
        Tool(
            name="list_s3_objects",
            description="List objects in the SAP BTP Object Store (S3) bucket.",
            inputSchema={
                "type": "object",
                "properties": {
                    "prefix":   {"type": "string",  "description": "Filter by key prefix (optional)"},
                    "max_keys": {"type": "integer", "description": "Max results (default: 100)", "default": 100},
                },
                "required": [],
            },
        ),
        Tool(
            name="get_s3_presigned_url",
            description="Generate a presigned download URL for an S3 object. Expires after 1 hour.",
            inputSchema={
                "type": "object",
                "properties": {
                    "s3_key": {"type": "string", "description": "S3 object key"},
                },
                "required": ["s3_key"],
            },
        ),
    ]


# ── Tool handlers ─────────────────────────────────────────────────────────────

@_mcp.call_tool()
async def _call_tool(name: str, arguments: dict):
    try:
        if name == "search_documents":
            # Deferred import avoids circular dependency at module load time
            from api import _get_embed_model
            query  = arguments.get("query", "")
            top_k  = int(arguments.get("top_k", 5))
            model  = _get_embed_model()
            vector = model.encode(query).tolist()
            results = db.query(vector, top_k)
            return [TextContent(type="text", text=json.dumps(results, indent=2, ensure_ascii=False))]

        elif name == "list_indexed_files":
            files = db.get_files()
            return [TextContent(type="text", text=json.dumps(files, indent=2, ensure_ascii=False))]

        elif name == "get_db_stats":
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
            return [TextContent(type="text", text=json.dumps(summary, indent=2, ensure_ascii=False))]

        elif name == "list_s3_objects":
            prefix   = arguments.get("prefix", "")
            max_keys = int(arguments.get("max_keys", 100))
            objects  = s3.list_objects(prefix=prefix, max_keys=max_keys)
            return [TextContent(type="text", text=json.dumps(objects, indent=2, ensure_ascii=False))]

        elif name == "get_s3_presigned_url":
            s3_key = arguments.get("s3_key", "")
            url    = s3.get_presigned_url(s3_key, expiry=3600)
            return [TextContent(type="text", text=json.dumps({"url": url, "expires_in": "1 hour"}))]

        else:
            return [TextContent(type="text", text=f"Unknown tool: {name}")]

    except Exception as e:
        log.error(f"MCP tool '{name}' error: {e}", exc_info=True)
        return [TextContent(type="text", text=f"Error: {e}")]


# ── FastAPI router ─────────────────────────────────────────────────────────────

router = APIRouter(tags=["mcp"])


@router.get("/mcp/sse", include_in_schema=False)
async def mcp_sse(request: Request):
    """SSE stream — Claude Desktop connects here to discover tools."""
    async with _sse.connect_sse(request.scope, request.receive, request._send) as streams:
        await _mcp.run(streams[0], streams[1], _mcp.create_initialization_options())


@router.post("/mcp/messages/", include_in_schema=False)
async def mcp_messages(request: Request):
    """Message handler — Claude Desktop posts tool call requests here."""
    await _sse.handle_post_message(request.scope, request.receive, request._send)
