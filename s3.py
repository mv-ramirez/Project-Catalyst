"""
s3.py — S3 client utilities for the SharePoint RAG API.

Credential resolution (highest priority first):
  1. VCAP_SERVICES env var  — SAP BTP Object Store / aws-s3 bound service
  2. s3_creds.json file     — local development override
  3. Environment variables  — AWS_ACCESS_KEY_ID / AWS_SECRET_ACCESS_KEY / …

All public functions are stateless wrappers around a module-level singleton
boto3 client so the TCP connection is reused across calls.
"""

from __future__ import annotations

import io
import json
import logging
import os
import threading
from typing import Any, Iterator

import boto3
from botocore.exceptions import ClientError

log = logging.getLogger("s3_client")

# ── Module-level singletons ───────────────────────────────────────────────────

_client: Any          = None
_bucket: str          = ""
_client_lock          = threading.Lock()


# ── Credential resolution ─────────────────────────────────────────────────────

# SAP BTP Object Store service plan names (expand as needed)
_OBJECTSTORE_SERVICE_KEYS = (
    "objectstore",
    "object-store",
    "aws-s3",
    "s3",
    "s3-compatible",
)


def _get_creds() -> dict:
    """
    Return a dict with keys:
        access_key_id, secret_access_key, bucket, region,
        endpoint_url (optional), host (optional)
    """

    # 1 ── VCAP_SERVICES (CF / SAP BTP bound service) ─────────────────────────
    vcap_raw = os.environ.get("VCAP_SERVICES")
    if vcap_raw:
        try:
            services = json.loads(vcap_raw)
        except json.JSONDecodeError as e:
            log.error("VCAP_SERVICES is present but not valid JSON: %s", e, exc_info=True)
            services = {}

        log.info("VCAP_SERVICES top-level keys: %s", list(services.keys()))

        # Try known Object Store service plan keys first
        for key in _OBJECTSTORE_SERVICE_KEYS:
            if key in services:
                creds = services[key][0]["credentials"]
                log.info("S3 credentials loaded from VCAP_SERVICES key '%s'", key)
                return _normalize_vcap_creds(creds)

        log.warning(
            "None of the known objectstore keys %s matched VCAP_SERVICES — "
            "falling back to full scan. Available keys: %s",
            _OBJECTSTORE_SERVICE_KEYS, list(services.keys())
        )

        # Fall back: scan all bound services for one that looks like S3
        # Handles both snake_case (access_key_id) and camelCase (accessKeyId)
        for svc_label, instances in services.items():
            for inst in instances:
                creds = inst.get("credentials", {})
                has_key = "access_key_id" in creds or "accessKeyId" in creds
                has_secret = "secret_access_key" in creds or "secretAccessKey" in creds
                if has_key and has_secret:
                    log.info(
                        "S3 credentials found via full scan — VCAP label='%s' instance='%s'  cred_keys=%s",
                        svc_label, inst.get("name", "unknown"), list(creds.keys()),
                    )
                    return _normalize_vcap_creds(creds)

        log.warning("VCAP_SERVICES present but no S3-compatible credentials found in any service binding")

    # 2 ── s3_creds.json (local development) ──────────────────────────────────
    creds_file = os.path.join(os.path.dirname(__file__), "s3_creds.json")
    if os.path.isfile(creds_file):
        with open(creds_file, encoding="utf-8") as fh:
            data = json.load(fh)
        if data.get("access_key_id") and data["access_key_id"] != "YOUR_ACCESS_KEY_ID":
            log.info("S3 credentials loaded from s3_creds.json")
            return {
                "access_key_id":     data["access_key_id"],
                "secret_access_key": data["secret_access_key"],
                "bucket":            data.get("bucket", ""),
                "region":            data.get("region", "us-east-1"),
                "endpoint_url":      data.get("endpoint_url", ""),
            }

    # 3 ── Environment variables (fallback / CI) ───────────────────────────────
    log.info("S3 credentials loaded from environment variables")
    return {
        "access_key_id":     os.getenv("AWS_ACCESS_KEY_ID", ""),
        "secret_access_key": os.getenv("AWS_SECRET_ACCESS_KEY", ""),
        "bucket":            os.getenv("S3_BUCKET", ""),
        "region":            os.getenv("AWS_REGION", os.getenv("AWS_DEFAULT_REGION", "us-east-1")),
        "endpoint_url":      os.getenv("S3_ENDPOINT_URL", ""),
    }


def _normalize_vcap_creds(raw: dict) -> dict:
    """Map SAP BTP Object Store credential keys to our standard shape."""
    endpoint = raw.get("endpoint_url", raw.get("host", ""))
    # boto3 endpoint_url must be a full URL — SAP BTP 'host' is a bare hostname
    if endpoint and not endpoint.startswith(("http://", "https://")):
        endpoint = f"https://{endpoint}"
    return {
        "access_key_id":     raw.get("access_key_id", raw.get("accessKeyId", "")),
        "secret_access_key": raw.get("secret_access_key", raw.get("secretAccessKey", "")),
        "bucket":            raw.get("bucket", raw.get("bucket_name", "")),
        "region":            raw.get("region", raw.get("aws_region", "us-east-1")),
        "endpoint_url":      endpoint,
    }


# ── Client factory ────────────────────────────────────────────────────────────

def _get_client() -> tuple[Any, str]:
    """Return (boto3 S3 client, bucket_name), creating the client once."""
    global _client, _bucket
    with _client_lock:
        if _client is not None:
            return _client, _bucket

        creds = _get_creds()
        if not creds.get("access_key_id"):
            raise RuntimeError(
                "S3 credentials not found. Set VCAP_SERVICES, s3_creds.json, "
                "or AWS_ACCESS_KEY_ID / AWS_SECRET_ACCESS_KEY env vars."
            )

        boto_kwargs: dict = {
            "aws_access_key_id":     creds["access_key_id"],
            "aws_secret_access_key": creds["secret_access_key"],
            "region_name":           creds.get("region", "us-east-1"),
        }
        if creds.get("endpoint_url"):
            boto_kwargs["endpoint_url"] = creds["endpoint_url"]

        _client = boto3.client("s3", **boto_kwargs)
        _bucket  = creds.get("bucket", "")

        log.info(
            "S3 client initialised — bucket=%s region=%s endpoint=%s",
            _bucket,
            creds.get("region"),
            creds.get("endpoint_url") or "(AWS default)",
        )
        return _client, _bucket


def reset_client() -> None:
    """Force re-initialisation of the client (useful after credential rotation)."""
    global _client, _bucket
    with _client_lock:
        _client = None
        _bucket  = ""


# ── Public helpers ────────────────────────────────────────────────────────────

def get_bucket_name() -> str:
    """Return the configured bucket name."""
    _, bucket = _get_client()
    return bucket


# ── Object operations ─────────────────────────────────────────────────────────

def list_objects(
    prefix: str = "",
    delimiter: str = "",
    max_keys: int = 1000,
) -> list[dict]:
    """
    List objects in the bucket.

    Returns a list of dicts with keys:
        key, size, last_modified, etag, storage_class
    """
    client, bucket = _get_client()
    kwargs: dict = {"Bucket": bucket, "MaxKeys": max_keys}
    if prefix:
        kwargs["Prefix"] = prefix
    if delimiter:
        kwargs["Delimiter"] = delimiter

    results: list[dict] = []
    paginator = client.get_paginator("list_objects_v2")
    for page in paginator.paginate(**kwargs):
        for obj in page.get("Contents", []):
            results.append({
                "key":           obj["Key"],
                "size":          obj["Size"],
                "last_modified": obj["LastModified"].isoformat(),
                "etag":          obj.get("ETag", "").strip('"'),
                "storage_class": obj.get("StorageClass", "STANDARD"),
            })
    return results


def object_exists(s3_key: str) -> bool:
    """Return True if the object exists in the bucket."""
    client, bucket = _get_client()
    try:
        client.head_object(Bucket=bucket, Key=s3_key)
        return True
    except ClientError as e:
        if e.response["Error"]["Code"] in ("404", "NoSuchKey"):
            return False
        raise


def get_metadata(s3_key: str) -> dict:
    """
    Return object metadata (head_object).

    Keys: content_type, content_length, last_modified, etag, metadata
    """
    client, bucket = _get_client()
    resp = client.head_object(Bucket=bucket, Key=s3_key)
    return {
        "content_type":   resp.get("ContentType", ""),
        "content_length": resp.get("ContentLength", 0),
        "last_modified":  resp["LastModified"].isoformat(),
        "etag":           resp.get("ETag", "").strip('"'),
        "metadata":       resp.get("Metadata", {}),
    }


# ── Upload ────────────────────────────────────────────────────────────────────

def upload_file(
    local_path: str,
    s3_key: str,
    content_type: str = "application/octet-stream",
    extra_args: dict | None = None,
) -> str:
    """
    Upload a local file to S3.

    Returns the s3_key of the uploaded object.
    """
    client, bucket = _get_client()
    args: dict = {"ContentType": content_type}
    if extra_args:
        args.update(extra_args)

    client.upload_file(
        Filename=local_path,
        Bucket=bucket,
        Key=s3_key,
        ExtraArgs=args,
    )
    log.info("Uploaded file '%s' → s3://%s/%s", local_path, bucket, s3_key)
    return s3_key


def upload_bytes(
    data: bytes | str,
    s3_key: str,
    content_type: str = "application/octet-stream",
    metadata: dict | None = None,
) -> str:
    """
    Upload in-memory bytes (or a UTF-8 string) to S3.

    Returns the s3_key of the uploaded object.
    """
    client, bucket = _get_client()
    if isinstance(data, str):
        data = data.encode("utf-8")

    put_kwargs: dict = {
        "Bucket":      bucket,
        "Key":         s3_key,
        "Body":        data,
        "ContentType": content_type,
    }
    if metadata:
        put_kwargs["Metadata"] = {k: str(v) for k, v in metadata.items()}

    client.put_object(**put_kwargs)
    log.info("Uploaded %d bytes → s3://%s/%s", len(data), bucket, s3_key)
    return s3_key


def upload_fileobj(
    fileobj: io.IOBase,
    s3_key: str,
    content_type: str = "application/octet-stream",
) -> str:
    """
    Upload a file-like object (e.g. an open file handle or BytesIO) to S3.

    Returns the s3_key of the uploaded object.
    """
    client, bucket = _get_client()
    client.upload_fileobj(
        fileobj,
        bucket,
        s3_key,
        ExtraArgs={"ContentType": content_type},
    )
    log.info("Uploaded fileobj → s3://%s/%s", bucket, s3_key)
    return s3_key


# ── Download ──────────────────────────────────────────────────────────────────

def download_file(s3_key: str, local_path: str) -> str:
    """
    Download an S3 object to a local file.

    Returns the local_path written.
    """
    client, bucket = _get_client()
    os.makedirs(os.path.dirname(local_path) or ".", exist_ok=True)
    client.download_file(Bucket=bucket, Key=s3_key, Filename=local_path)
    log.info("Downloaded s3://%s/%s → %s", bucket, s3_key, local_path)
    return local_path


def download_bytes(s3_key: str) -> bytes:
    """Download an S3 object and return its raw bytes."""
    client, bucket = _get_client()
    resp = client.get_object(Bucket=bucket, Key=s3_key)
    data = resp["Body"].read()
    log.info("Downloaded s3://%s/%s (%d bytes)", bucket, s3_key, len(data))
    return data


def download_text(s3_key: str, encoding: str = "utf-8") -> str:
    """Download an S3 object and return it as a decoded string."""
    return download_bytes(s3_key).decode(encoding)


def stream_object(s3_key: str, chunk_size: int = 8192) -> Iterator[bytes]:
    """
    Yield chunks of an S3 object without loading it all into memory.
    Useful for streaming large files through FastAPI StreamingResponse.
    """
    client, bucket = _get_client()
    resp   = client.get_object(Bucket=bucket, Key=s3_key)
    stream = resp["Body"]
    while True:
        chunk = stream.read(chunk_size)
        if not chunk:
            break
        yield chunk


# ── Delete ────────────────────────────────────────────────────────────────────

def delete_object(s3_key: str) -> None:
    """Delete a single object from the bucket."""
    client, bucket = _get_client()
    client.delete_object(Bucket=bucket, Key=s3_key)
    log.info("Deleted s3://%s/%s", bucket, s3_key)


def delete_objects(s3_keys: list[str]) -> dict:
    """
    Bulk-delete up to 1 000 objects in a single API call.

    Returns the raw S3 response (contains 'Deleted' and 'Errors' lists).
    """
    if not s3_keys:
        return {"Deleted": [], "Errors": []}
    client, bucket = _get_client()
    resp = client.delete_objects(
        Bucket=bucket,
        Delete={"Objects": [{"Key": k} for k in s3_keys], "Quiet": False},
    )
    deleted = resp.get("Deleted", [])
    errors  = resp.get("Errors",  [])
    log.info("Bulk delete: %d deleted, %d errors", len(deleted), len(errors))
    return {"Deleted": deleted, "Errors": errors}


def delete_prefix(prefix: str) -> int:
    """
    Delete all objects whose key starts with *prefix*.

    Returns the number of objects deleted.
    """
    keys = [o["key"] for o in list_objects(prefix=prefix)]
    if not keys:
        return 0
    # boto3 delete_objects supports max 1000 per call — batch if needed
    total = 0
    for i in range(0, len(keys), 1000):
        result = delete_objects(keys[i : i + 1000])
        total += len(result["Deleted"])
    return total


# ── Copy ──────────────────────────────────────────────────────────────────────

def copy_object(
    src_key: str,
    dest_key: str,
    src_bucket: str | None = None,
) -> None:
    """
    Copy an object within (or across) buckets.

    If src_bucket is None the configured bucket is used as source.
    """
    client, bucket = _get_client()
    copy_source = {"Bucket": src_bucket or bucket, "Key": src_key}
    client.copy_object(CopySource=copy_source, Bucket=bucket, Key=dest_key)
    log.info(
        "Copied s3://%s/%s → s3://%s/%s",
        src_bucket or bucket, src_key, bucket, dest_key,
    )


# ── Presigned URLs ────────────────────────────────────────────────────────────

def get_presigned_url(s3_key: str, expiry: int = 3600) -> str:
    """
    Generate a presigned GET URL that allows anyone to download the object
    without AWS credentials.  Expires after *expiry* seconds (default 1 hour).
    """
    client, bucket = _get_client()
    url = client.generate_presigned_url(
        "get_object",
        Params={"Bucket": bucket, "Key": s3_key},
        ExpiresIn=expiry,
    )
    log.info("Presigned GET URL for '%s' (expires in %ds)", s3_key, expiry)
    return url


def get_presigned_upload_url(
    s3_key: str,
    expiry: int = 3600,
    content_type: str = "application/octet-stream",
    max_size_bytes: int = 100 * 1024 * 1024,
) -> dict:
    """
    Generate a presigned POST policy for direct browser/client uploads.

    Returns a dict with 'url' and 'fields' — pass both to the HTTP client.
    max_size_bytes defaults to 100 MB.
    """
    client, bucket = _get_client()
    resp = client.generate_presigned_post(
        Bucket=bucket,
        Key=s3_key,
        Fields={"Content-Type": content_type},
        Conditions=[
            {"Content-Type": content_type},
            ["content-length-range", 1, max_size_bytes],
        ],
        ExpiresIn=expiry,
    )
    log.info("Presigned POST URL for '%s' (expires in %ds)", s3_key, expiry)
    return resp  # {"url": "...", "fields": {...}}


# ── Health check ──────────────────────────────────────────────────────────────

def ping() -> dict:
    """
    Verify connectivity by fetching bucket metadata.

    Returns {"ok": True, "bucket": "...", "region": "..."}
    Raises RuntimeError on failure.
    """
    client, bucket = _get_client()
    try:
        resp = client.head_bucket(Bucket=bucket)
        region = resp.get("ResponseMetadata", {}).get(
            "HTTPHeaders", {}
        ).get("x-amz-bucket-region", "unknown")
        return {"ok": True, "bucket": bucket, "region": region}
    except ClientError as e:
        code = e.response["Error"]["Code"]
        raise RuntimeError(f"S3 ping failed — bucket={bucket} error={code}") from e


# ── Standalone test ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    print("=== S3 connectivity test ===")
    try:
        result = ping()
        print(f"[OK] Connected — bucket={result['bucket']} region={result['region']}")
    except Exception as exc:
        print(f"[FAIL] {exc}")
        sys.exit(1)

    print("\n=== List objects (first 10) ===")
    objects = list_objects(max_keys=10)
    if objects:
        for obj in objects:
            print(f"  {obj['key']}  ({obj['size']} bytes)  {obj['last_modified']}")
    else:
        print("  (bucket is empty)")

    print("\n=== Round-trip: upload → download → delete ===")
    test_key = "__s3_test/roundtrip.txt"
    payload  = "Hello from s3.py round-trip test"
    upload_bytes(payload, test_key, content_type="text/plain")
    retrieved = download_text(test_key)
    assert retrieved == payload, f"Mismatch: {retrieved!r}"
    delete_object(test_key)
    print(f"  [OK] key={test_key!r} passed")

    print("\nAll tests passed.")
