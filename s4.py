"""
s4.py — SAP S/4HANA ABAP Development Tools (ADT) client

Publishes ABAP source files to S/4HANA Cloud via the ADT REST API.

Auth priority (in order):
  1. S4_USER / S4_PASSWORD env vars  (Communication User — bypasses SSO)
  2. BTP Destination Service authToken  (BTP-Public-Cloud → BTP_USER_001 via cert)
  3. credentials passed explicitly in the API request body

The BTP-Public-Cloud destination uses ClientCertificateAuthentication.
The Destination Service resolves this to a pre-built Authorization header
(type=Basic, user=BTP_USER_001) that bypasses SAML entirely.
BTP_USER_001 must have the ADT authorization role (SAP_BC_DWB_ABAPDEVELOPER)
assigned in S/4HANA for the publish flow to succeed.
"""

import os
import re
import logging
import requests

log = logging.getLogger("s4_adt")

S4_BASE_URL    = os.getenv("S4_BASE_URL",    "https://my407593.s4hana.cloud.sap")
S4_CLIENT      = os.getenv("S4_CLIENT",      "080")
S4_DESTINATION = os.getenv("S4_DESTINATION", "BTP-Public-Cloud")

ADT_PROGRAMS  = "/sap/bc/adt/programs/programs"
ADT_DISCOVERY = "/sap/bc/adt/discovery"


# ── Credential resolution ─────────────────────────────────────────────────────

def get_s4_session(dest_svc_token: str | None = None,
                   dest_svc_uri: str | None = None,
                   override_user: str | None = None,
                   override_pass: str | None = None) -> tuple[requests.Session, str, str]:
    """
    Returns (session, base_url, sap_client) with auth configured.

    Priority:
      1. override_user / override_pass  (from request body — explicit credentials)
      2. S4_USER / S4_PASSWORD env vars (Communication User)
      3. BTP Destination Service authToken for BTP-Public-Cloud
         → Destination Service resolves the client cert to BTP_USER_001 and
           returns a ready-to-use Authorization: Basic header that bypasses SAML.
    """
    base_url = S4_BASE_URL.rstrip("/")
    client   = S4_CLIENT

    # Option 1 — explicit credentials from caller
    if override_user and override_pass:
        log.info("S4: using caller-supplied credentials")
        return _session_from_userpass(override_user, override_pass, client), base_url, client

    # Option 2 — env var Communication User
    env_user = os.getenv("S4_USER")
    env_pass = os.getenv("S4_PASSWORD")
    if env_user and env_pass:
        log.info(f"S4: using env-var Communication User '{env_user}'")
        return _session_from_userpass(env_user, env_pass, client), base_url, client

    # Option 3 — BTP Destination Service (BTP-Public-Cloud → BTP_USER_001)
    if dest_svc_token and dest_svc_uri:
        log.info(f"S4: resolving auth via BTP Destination '{S4_DESTINATION}'")
        r = requests.get(
            f"{dest_svc_uri}/destination-configuration/v1/destinations/{S4_DESTINATION}",
            headers={
                "Authorization": f"Bearer {dest_svc_token}",
                "X-user-token":  dest_svc_token,
            },
            timeout=15,
        )
        r.raise_for_status()
        payload    = r.json()
        config     = payload.get("destinationConfiguration", {})
        dest_url   = config.get("URL", base_url).rstrip("/")
        dest_cli   = config.get("sap-client", client)
        auth_token = next((t for t in payload.get("authTokens", []) if not t.get("error")), None)

        if auth_token:
            token_type  = auth_token.get("type", "Basic")
            token_value = auth_token.get("value", "")
            log.info(f"S4: authToken type={token_type} from Destination Service (user=BTP_USER_001)")
            return _session_from_token(token_type, token_value, dest_cli), dest_url, dest_cli

        # Fallback — try plain user/password fields if present
        dest_user = config.get("User") or config.get("user", "")
        dest_pass = config.get("Password") or config.get("password", "")
        if dest_user and dest_pass:
            log.info(f"S4: falling back to destination user='{dest_user}'")
            return _session_from_userpass(dest_user, dest_pass, dest_cli), dest_url, dest_cli

    raise RuntimeError(
        "No S/4HANA credentials available. "
        "Running on CF: ensure the BTP Destination Service is bound and "
        f"destination '{S4_DESTINATION}' is configured. "
        "Local: set S4_USER + S4_PASSWORD env vars (Communication User), "
        "or pass credentials in the request body."
    )


def _session_from_userpass(user: str, password: str, client: str) -> requests.Session:
    s = requests.Session()
    s.auth = (user, password)
    s.headers.update({"sap-client": client})
    return s


def _session_from_token(token_type: str, token_value: str, client: str) -> requests.Session:
    """Build a session using a pre-built auth token (e.g. Basic <base64>) from Destination Service."""
    s = requests.Session()
    s.headers.update({
        "Authorization": f"{token_type} {token_value}",
        "sap-client":    client,
    })
    return s


# ── ADT operations ────────────────────────────────────────────────────────────

def fetch_csrf_token(session: requests.Session, base_url: str) -> str:
    """Fetch ADT CSRF token. Raises descriptive errors on auth/authz failure."""
    r = session.get(
        f"{base_url}{ADT_DISCOVERY}",
        headers={"Accept": "application/xml", "X-CSRF-Token": "Fetch"},
        timeout=20,
    )

    if r.text.strip().startswith("<html"):
        raise RuntimeError(
            "S/4HANA redirected to SAML SSO — the user resolved by the destination "
            "is not accepted for programmatic ADT access. "
            "Check that BTP-Public-Cloud destination is configured and BTP_USER_001 exists."
        )

    if r.status_code == 401:
        raise RuntimeError("S/4HANA returned 401 — credentials rejected.")

    if r.status_code == 403:
        raise RuntimeError(
            "S/4HANA returned 403 — BTP_USER_001 authenticated but lacks ADT authorization. "
            "Ask the S/4HANA admin to assign role SAP_BC_DWB_ABAPDEVELOPER to BTP_USER_001."
        )

    if r.status_code != 200:
        raise RuntimeError(f"ADT discovery returned {r.status_code}: {r.text[:300]}")

    csrf = r.headers.get("X-CSRF-Token")
    if not csrf:
        raise RuntimeError("ADT discovery succeeded but no X-CSRF-Token returned.")

    log.info("S4: CSRF token fetched successfully")
    return csrf


def program_exists(session: requests.Session, base_url: str, program_name: str) -> bool:
    r = session.get(
        f"{base_url}{ADT_PROGRAMS}/{program_name}",
        headers={"Accept": "application/xml"},
        timeout=15,
    )
    return r.status_code == 200


def create_program(session: requests.Session, base_url: str, csrf_token: str,
                   package: str, program_name: str, description: str = "") -> None:
    xml_body = f"""<?xml version="1.0" encoding="utf-8"?>
<program:abapProgram
    xmlns:program="http://www.sap.com/adt/programs"
    xmlns:adtcore="http://www.sap.com/adt/core"
    adtcore:description="{description or program_name}"
    adtcore:name="{program_name}"
    adtcore:masterLanguage="EN"
    adtcore:responsible=""
    adtcore:packageName="{package}">
</program:abapProgram>"""

    r = session.post(
        f"{base_url}{ADT_PROGRAMS}",
        data=xml_body.encode("utf-8"),
        headers={
            "Content-Type": "application/vnd.sap.adt.programs+xml; charset=utf-8",
            "Accept":       "application/xml",
            "X-CSRF-Token": csrf_token,
        },
        timeout=30,
    )
    if r.status_code not in (200, 201):
        raise RuntimeError(
            f"Failed to create ABAP program '{program_name}': {r.status_code} — {r.text[:400]}"
        )
    log.info(f"S4: program '{program_name}' created in package '{package}'")


def lock_program(session: requests.Session, base_url: str,
                 csrf_token: str, program_name: str) -> str:
    r = session.post(
        f"{base_url}{ADT_PROGRAMS}/{program_name}",
        params={"_action": "LOCK", "accessMode": "MODIFY"},
        headers={
            "Accept":       "application/vnd.sap.adt.lock+xml; charset=utf-8",
            "X-CSRF-Token": csrf_token,
        },
        timeout=15,
    )
    if r.status_code != 200:
        raise RuntimeError(f"Failed to lock '{program_name}': {r.status_code} — {r.text[:300]}")

    match = re.search(r"<adtcore:lockHandle[^>]*>([^<]+)</adtcore:lockHandle>", r.text)
    if not match:
        try:
            lock_handle = r.json().get("lockHandle", "")
            if lock_handle:
                return lock_handle
        except Exception:
            pass
        raise RuntimeError(f"Lock response did not contain lockHandle: {r.text[:300]}")

    lock_handle = match.group(1)
    log.info(f"S4: '{program_name}' locked (handle={lock_handle[:12]}…)")
    return lock_handle


def update_source(session: requests.Session, base_url: str, csrf_token: str,
                  program_name: str, lock_handle: str, source_code: str) -> None:
    r = session.put(
        f"{base_url}{ADT_PROGRAMS}/{program_name}/source/main",
        data=source_code.encode("utf-8"),
        params={"lockHandle": lock_handle},
        headers={
            "Content-Type": "text/plain; charset=utf-8",
            "Accept":       "text/plain",
            "X-CSRF-Token": csrf_token,
        },
        timeout=60,
    )
    if r.status_code not in (200, 204):
        raise RuntimeError(
            f"Failed to upload source for '{program_name}': {r.status_code} — {r.text[:400]}"
        )
    log.info(f"S4: source uploaded for '{program_name}' ({len(source_code)} chars)")


def unlock_program(session: requests.Session, base_url: str,
                   csrf_token: str, program_name: str, lock_handle: str) -> None:
    session.delete(
        f"{base_url}{ADT_PROGRAMS}/{program_name}",
        params={"_action": "UNLOCK", "lockHandle": lock_handle},
        headers={"X-CSRF-Token": csrf_token},
        timeout=15,
    )
    log.info(f"S4: '{program_name}' unlocked")


def publish_abap(session: requests.Session, base_url: str,
                 package: str, program_name: str, source_code: str,
                 description: str = "") -> dict:
    """
    Full publish flow:
      1. Fetch CSRF token (validates auth + authz)
      2. Create program if it doesn't exist
      3. Lock → upload source → unlock
    """
    csrf   = fetch_csrf_token(session, base_url)
    exists = program_exists(session, base_url, program_name)
    if not exists:
        create_program(session, base_url, csrf, package, program_name, description)

    lock_handle = lock_program(session, base_url, csrf, program_name)
    try:
        update_source(session, base_url, csrf, program_name, lock_handle, source_code)
    finally:
        unlock_program(session, base_url, csrf, program_name, lock_handle)

    return {
        "program":     program_name,
        "package":     package,
        "action":      "updated" if exists else "created",
        "source_size": len(source_code),
        "base_url":    base_url,
    }
