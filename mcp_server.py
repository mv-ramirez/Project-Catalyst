"""
MCP (Model Context Protocol) server — SSE transport for Claude Desktop.

Uses mcp v1.x API (pinned mcp>=1.0.0,<2.0.0). Tools exposed:
  search_documents      — semantic vector search across HANA
  list_indexed_files    — all files in BRAIN_SP_FILES
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
        Tool(
            name="publish_abap",
            description=(
                "Download an ABAP source file from S3 and publish it to S/4HANA Cloud via ADT REST API. "
                "Requires S4_USER and S4_PASSWORD to be configured in the BTP environment, "
                "or pass s4_user/s4_password in the arguments. "
                "The Communication User must have the ADT business catalog assigned."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "s3_key":       {"type": "string", "description": "S3 object key of the .abap source file"},
                    "package":      {"type": "string", "description": "ABAP package (e.g. ZLOCAL)"},
                    "program_name": {"type": "string", "description": "ABAP program name (e.g. ZMYPROGRAM)"},
                    "s4_user":      {"type": "string", "description": "Override Communication User (optional)"},
                    "s4_password":  {"type": "string", "description": "Override Communication User password (optional)"},
                },
                "required": ["s3_key", "package", "program_name"],
            },
        ),
        Tool(
            name="test_s4_connection",
            description="Test connectivity and authentication to S/4HANA ADT. Returns connection status and CSRF token validity.",
            inputSchema={
                "type": "object",
                "properties": {
                    "s4_user":     {"type": "string", "description": "Override Communication User (optional)"},
                    "s4_password": {"type": "string", "description": "Override password (optional)"},
                },
                "required": [],
            },
        ),
    ]


# ── Tool handlers ─────────────────────────────────────────────────────────────

@_mcp.call_tool()
async def _call_tool(name: str, arguments: dict):
    try:
        if name == "search_documents":
            from api import _get_embed_model  # deferred — avoids circular import at load time
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

        elif name == "publish_abap":
            import s4
            s3_key       = arguments.get("s3_key", "")
            package      = arguments.get("package", "")
            program_name = arguments.get("program_name", "")
            s4_user      = arguments.get("s4_user")
            s4_password  = arguments.get("s4_password")

            # Download source from S3
            source_bytes = s3.download_bytes(s3_key)
            source_code  = source_bytes.decode("utf-8")

            # Publish to S/4HANA via ADT
            session, base_url, _ = s4.get_s4_session(
                override_user=s4_user,
                override_pass=s4_password,
            )
            result = s4.publish_abap(
                session=session,
                base_url=base_url,
                package=package,
                program_name=program_name,
                source_code=source_code,
            )
            return [TextContent(type="text", text=json.dumps(result, indent=2, ensure_ascii=False))]

        elif name == "test_s4_connection":
            import s4
            s4_user     = arguments.get("s4_user")
            s4_password = arguments.get("s4_password")
            try:
                session, base_url, _ = s4.get_s4_session(
                    override_user=s4_user,
                    override_pass=s4_password,
                )
                csrf = s4.fetch_csrf_token(session, base_url)
                result = {
                    "connected":   True,
                    "base_url":    base_url,
                    "csrf_token":  csrf[:8] + "..." if csrf else None,
                    "auth_method": "override" if s4_user else ("env_var" if __import__("os").getenv("S4_USER") else "destination"),
                }
            except Exception as exc:
                result = {"connected": False, "error": str(exc), "base_url": s4.S4_BASE_URL}
            return [TextContent(type="text", text=json.dumps(result, indent=2, ensure_ascii=False))]

        else:
            return [TextContent(type="text", text=f"Unknown tool: {name}")]

    except Exception as e:
        log.error(f"MCP tool '{name}' error: {e}", exc_info=True)
        return [TextContent(type="text", text=f"Error: {e}")]


# ── FastAPI router ─────────────────────────────────────────────────────────────

router = APIRouter(tags=["mcp"])


@router.get("/mcp/sse", include_in_schema=False)
async def mcp_sse(request: Request):
    """SSE stream — Claude Desktop connects here to discover and call tools."""
    async with _sse.connect_sse(request.scope, request.receive, request._send) as streams:
        await _mcp.run(streams[0], streams[1], _mcp.create_initialization_options())


@router.post("/mcp/messages/", include_in_schema=False)
async def mcp_messages(request: Request):
    """Message handler — Claude Desktop posts tool call requests here."""
    await _sse.handle_post_message(request.scope, request.receive, request._send)
