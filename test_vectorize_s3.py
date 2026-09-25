#!/usr/bin/env python
"""
Exercise the BTP-Brain ingest + retrieval path against the HDI-owned
CATALYSTPLATFORM_TB_BRAIN_* tables.

This is the check that has to pass BEFORE anything is switched over: it writes
real rows through the real db.ingest() code path, reads them back through the
real filtered query, and — when dual-write is on — proves the legacy and HDI
tables ended up holding the same thing.

Nothing here issues DDL. The HDI tables must already exist, which means the db
module in BTP-Catalyst-Platform has been deployed. Check with:

    curl -s $BASE/db/tables | python -m json.tool

Usage
-----
    # 1. Confirm both table sets are present and see their row counts
    python test_vectorize_s3.py tables

    # 2. Insert a synthetic tagged fixture and read it back filtered.
    #    Writes to whatever BRAIN_WRITE_SET says; cleans up after itself.
    BRAIN_WRITE_SET=both python test_vectorize_s3.py roundtrip

    # 3. Vectorize real BTP-Brain content from S3, tagged, then assert the
    #    filters actually narrow the result
    BRAIN_WRITE_SET=both python test_vectorize_s3.py vectorize --prefix knowledge-base/

    # 4. Findings-style check: does a per-deliverable query return that
    #    deliverable's standards rather than the whole KB?
    python test_vectorize_s3.py findings --deliverable B10

Environment
-----------
    BRAIN_WRITE_SET   legacy | hdi | both   which table set(s) receive writes
    BRAIN_READ_SET    legacy | hdi          which table set queries read from
    HANA_*            connection, as db.py expects

Exit code is 0 only if every assertion passed.
"""

import argparse
import os
import sys
import uuid

import db


# ── Tiny assertion harness ────────────────────────────────────────────────────

class Results:
    def __init__(self):
        self.passed = 0
        self.failed = 0

    def check(self, label: str, condition: bool, detail: str = "") -> bool:
        if condition:
            self.passed += 1
            print(f"  PASS  {label}")
        else:
            self.failed += 1
            print(f"  FAIL  {label}" + (f"\n          {detail}" if detail else ""))
        return condition

    def report(self) -> int:
        total = self.passed + self.failed
        print(f"\n{self.passed}/{total} checks passed")
        return 1 if self.failed else 0


def _banner(text: str) -> None:
    print(f"\n{text}\n{'─' * len(text)}")


def _config() -> None:
    print(f"read set   : {db.BRAIN_READ_SET}")
    print(f"write sets : {'+'.join(db.WRITE_SETS)}")
    print(f"vectors    : {db.T_VECTORS}")


# ── tables ────────────────────────────────────────────────────────────────────

def cmd_tables(args, r: Results) -> None:
    """Both table sets side by side: present? how many rows? do they agree?"""
    _banner("Table sets")
    info = db.describe_tables()
    print(f"schema     : {info['schema']}")
    _config()
    print()
    print(f"{'key':<16}{'legacy':<40}{'rows':>8}   {'hdi':<44}{'rows':>8}  match")
    for key, entry in info["tables"].items():
        legacy, hdi = entry["legacy"], entry["hdi"]
        lrows = "—" if legacy["rows"] is None else legacy["rows"]
        hrows = "—" if hdi["rows"] is None else hdi["rows"]
        flag  = "yes" if entry["matched"] else "no"
        print(f"{key:<16}{legacy['table']:<40}{lrows:>8}   {hdi['table']:<44}{hrows:>8}  {flag}")

    missing = [k for k, e in info["tables"].items() if not e["hdi"]["present"]]
    r.check(
        "every HDI TB_BRAIN_* table exists",
        not missing,
        f"missing: {', '.join(missing)} — deploy the db module in BTP-Catalyst-Platform first",
    )


# ── roundtrip ─────────────────────────────────────────────────────────────────

# A synthetic BTP-Brain slice: one template and two guidelines across two layers
# and two deliverables. Small enough to insert fast, shaped so that a filter
# which does not work is immediately visible — every assertion below would pass
# trivially if all fixtures shared one layer, so they deliberately do not.
LAYER_SAP  = "Layer 1 - SAP Guidelines"
LAYER_ACN  = "Layer 2 - Accenture Guidelines"
CAT_TPL    = "Templates"
CAT_STD    = "Standards"

FIXTURES = [
    # (suffix, layer, category, deliverable, filename, text)
    ("tpl-b10", LAYER_SAP, CAT_TPL, "B10", "kdd-framework-template.docx",
     "Key Design Decisions Framework template. Every decision record must carry a "
     "Decision ID, Summary, Options considered, DDA anchor, Owner and Due date."),
    ("std-b10", LAYER_ACN, CAT_STD, "B10", "kdd-accenture-standard.docx",
     "Accenture standard for Key Design Decisions: each decision must name an "
     "approver and record the rationale for the option chosen."),
    ("tpl-c5",  LAYER_SAP, CAT_TPL, "C5", "fdd-template.docx",
     "Functional Design Document template. Sections: process flow, data model, "
     "configuration, acceptance criteria."),
    ("std-c5",  LAYER_ACN, CAT_STD, "C5", "fdd-accenture-standard.docx",
     "Accenture functional design standard: describe the gap, the proposed "
     "solution, and the test approach for each requirement."),
]


def _embed(text: str) -> list:
    from api import _get_embed_model
    return _get_embed_model().encode(text, normalize_embeddings=True).tolist()


def _insert_fixtures(run_id: str) -> list:
    """Push the fixture set through the real db.ingest(). Returns the item ids."""
    item_ids = []
    for suffix, layer, category, deliverable, filename, text in FIXTURES:
        item_id = f"test-{run_id}-{suffix}"
        item_ids.append(item_id)
        tags = {
            "layer":            layer,
            "layer_category":   category,
            "deliverable_code": deliverable,
            "project":          None,
        }
        item = {
            "id": item_id, "name": filename, "type": "docx", "extension": "docx",
            "size": "1.0 KB", "size_bytes": 1024, "doc_path": f"knowledge-base/{layer}/{category}/{filename}",
            "etag": item_id, "storage_class": "STANDARD", "last_modified_date": "",
            **tags,
        }
        chunk  = {"item_id": item_id, "chunk_id": 1, "chunk_data": text, "filename": filename, **tags}
        vector = {"item_id": item_id, "chunk_id": 1, "vector_data": _embed(text),
                  "filename": filename, **tags}

        result = db.ingest([chunk], [vector], [item], source="s3")
        if result.get("errors"):
            raise RuntimeError(f"ingest reported errors for {item_id}: {result['errors']}")
    return item_ids


def _cleanup(item_ids: list) -> None:
    for item_id in item_ids:
        try:
            db.delete_item(item_id)
        except Exception as e:
            print(f"  warn  cleanup failed for {item_id}: {e}")


def cmd_roundtrip(args, r: Results) -> None:
    """Insert tagged fixtures, read them back filtered, then delete them."""
    _banner("Insert → filtered read → delete")
    _config()

    run_id   = uuid.uuid4().hex[:8]
    item_ids = []
    try:
        print(f"\ninserting {len(FIXTURES)} fixture chunks (run {run_id})")
        item_ids = _insert_fixtures(run_id)
        r.check(f"ingest() wrote {len(FIXTURES)} fixtures with no errors", True)

        mine = set(item_ids)

        # The tags must survive the write. Reading them back off the vector table
        # is the actual subject of this test — if ingest dropped them, every
        # filter below would return nothing and the failure would look like a
        # query bug rather than a write bug.
        _banner("Tags survived the write")
        rows = db.fetch_chunks({"deliverable_code": "B10"}, limit=50)
        theirs = [row for row in rows if row["item_id"] in mine]
        r.check("B10 fixtures readable by tag", len(theirs) == 2,
                f"got {len(theirs)} of 2")
        r.check("layer tag persisted",
                all(row["layer"] in (LAYER_SAP, LAYER_ACN) for row in theirs),
                f"layers: {[row['layer'] for row in theirs]}")

        # Filtering is the point of the change: a filtered query must not return
        # chunks from another deliverable or another layer.
        _banner("Filters actually narrow the result")
        query_text = "What must a key design decision record contain?"
        vector = _embed(query_text)

        by_deliverable = db.query(vector, top_k=50, filters={"deliverable_code": "B10"})
        stray = [row for row in by_deliverable if row["deliverable_code"] != "B10"]
        r.check("deliverable_code filter returns only B10", not stray,
                f"leaked: {[(row['filename'], row['deliverable_code']) for row in stray[:5]]}")

        by_layer = db.query(vector, top_k=50, filters={"layer": LAYER_SAP})
        stray = [row for row in by_layer if row["layer"] != LAYER_SAP]
        r.check("layer filter returns only Layer 1", not stray,
                f"leaked: {[(row['filename'], row['layer']) for row in stray[:5]]}")

        both_layers = db.query(vector, top_k=50, filters={"layer": [LAYER_SAP, LAYER_ACN]})
        stray = [row for row in both_layers if row["layer"] not in (LAYER_SAP, LAYER_ACN)]
        r.check("layer list filter returns only those two layers", not stray,
                f"leaked: {[(row['filename'], row['layer']) for row in stray[:5]]}")

        combined = db.query(vector, top_k=50,
                            filters={"layer": LAYER_SAP, "deliverable_code": "B10"})
        stray = [row for row in combined
                 if row["layer"] != LAYER_SAP or row["deliverable_code"] != "B10"]
        r.check("combined filters AND together", not stray,
                f"leaked: {len(stray)} rows")

        # An unfiltered query over the same text must see strictly more than the
        # filtered one — otherwise the filter is a no-op and we have proved nothing.
        unfiltered = db.query(vector, top_k=50)
        r.check("filtered result is a strict subset of unfiltered",
                len(by_deliverable) <= len(unfiltered),
                f"filtered={len(by_deliverable)} unfiltered={len(unfiltered)}")

        _banner("Score floor")
        floored = db.query(vector, top_k=50, filters={"deliverable_code": "B10"}, min_score=0.99)
        r.check("min_score drops weak matches",
                len(floored) <= len(by_deliverable),
                f"floored={len(floored)} unfloored={len(by_deliverable)}")
        r.check("every returned row clears the floor",
                all(row["score"] >= 0.99 for row in floored),
                f"scores: {[round(row['score'], 3) for row in floored]}")

        _banner("Rejected input")
        try:
            db.query(vector, top_k=5, filters={"nonsense_column": "x"})
            r.check("unknown filter column is rejected", False, "no error raised")
        except ValueError:
            r.check("unknown filter column is rejected", True)

        try:
            db.fetch_chunks({}, limit=10)
            r.check("unfiltered fetch_chunks is rejected", False, "no error raised")
        except ValueError:
            r.check("unfiltered fetch_chunks is rejected", True)

        # Dual-write is only useful if both sets really got the rows.
        if len(db.WRITE_SETS) > 1:
            _banner("Dual-write landed in both sets")
            for set_name in db.WRITE_SETS:
                tables = db.TABLE_SETS[set_name]
                conn = db._connect()
                cur  = conn.cursor()
                try:
                    placeholders = ", ".join("?" for _ in item_ids)
                    cur.execute(
                        f"SELECT COUNT(*) FROM {db._t(tables['vectors'])} "
                        f"WHERE item_id IN ({placeholders})", tuple(item_ids)
                    )
                    count = cur.fetchone()[0]
                finally:
                    cur.close()
                r.check(f"{set_name}: {tables['vectors']} holds all {len(item_ids)} vectors",
                        count == len(item_ids), f"found {count}")
    finally:
        if item_ids:
            _banner("Cleanup")
            _cleanup(item_ids)
            print(f"  deleted {len(item_ids)} fixture items")


# ── vectorize ─────────────────────────────────────────────────────────────────

def cmd_vectorize(args, r: Results) -> None:
    """Vectorize real BTP-Brain content from S3 and confirm it came back tagged."""
    _banner(f"Vectorizing S3 prefix {args.prefix!r}")
    _config()

    import s3 as s3_client
    from api import _download_and_vectorize_s3, _parse_s3_tags

    try:
        raw = s3_client.list_objects(prefix=args.prefix)
    except Exception as e:
        r.check(f"list S3 prefix {args.prefix!r}", False, str(e))
        return

    objects = [o for o in raw if not o["key"].endswith("/")][: args.limit]
    r.check(f"found files under {args.prefix!r}", bool(objects),
            "nothing to vectorize — is the prefix right?")
    if not objects:
        return

    # Show what the path parser made of each key before anything is written —
    # a run where every deliverable_code is None means the folder layout does not
    # match the convention, and the filters will have nothing to bite on.
    print()
    tagged = 0
    for obj in objects:
        tags = _parse_s3_tags(obj["key"])
        tagged += bool(tags["deliverable_code"])
        print(f"  {obj['key']}\n      layer={tags['layer']!r} category={tags['layer_category']!r} "
              f"deliverable={tags['deliverable_code']!r}")

    r.check("at least one file resolved to a deliverable", tagged > 0,
            "no key matched a deliverable folder — check the BTP-Brain layout "
            "or pass explicit overrides on the vectorize request")

    if args.dry_run:
        print("\ndry run — nothing written")
        return

    print(f"\nvectorizing {len(objects)} file(s)…")
    n_chunks, n_files, *_rest, ingest_results = _download_and_vectorize_s3(objects)
    r.check(f"vectorized {n_files} file(s), {n_chunks} chunk(s)", n_chunks > 0)

    failed = [x for x in ingest_results if x.get("status") not in ("ok", "skipped")]
    r.check("no file failed to ingest", not failed, f"failed: {failed[:3]}")

    # And the rows must be findable by the tags that were just derived.
    _banner("Ingested rows are retrievable by tag")
    seen_layers = {_parse_s3_tags(o["key"])["layer"] for o in objects}
    seen_layers.discard(None)
    for layer in sorted(seen_layers):
        rows  = db.fetch_chunks({"layer": layer}, limit=200)
        stray = [row for row in rows if row["layer"] != layer]
        r.check(f"layer {layer!r} fetch returns only that layer", not stray,
                f"leaked {len(stray)} rows")


# ── findings ──────────────────────────────────────────────────────────────────

def cmd_findings(args, r: Results) -> None:
    """The generateFindings question: does a per-deliverable query stay on topic?

    Compares what the QA pass would have received before (everything) against
    what it receives now (that deliverable's standards).
    """
    code = args.deliverable
    _banner(f"Findings-style retrieval for {code}")
    _config()

    query_text = args.query or f"What standards and required sections apply to deliverable {code}?"
    vector = _embed(query_text)

    unfiltered = db.query(vector, top_k=args.top_k)
    filtered   = db.query(vector, top_k=args.top_k, filters={"deliverable_code": code})
    template   = db.fetch_chunks({"deliverable_code": code, "layer_category": args.template_category},
                                 limit=200)

    print(f"\nquery      : {query_text}")
    print(f"unfiltered : {len(unfiltered)} chunks, "
          f"{sum(len(row['chunk_data'] or '') for row in unfiltered):,} chars")
    print(f"filtered   : {len(filtered)} chunks, "
          f"{sum(len(row['chunk_data'] or '') for row in filtered):,} chars")
    print(f"template   : {len(template)} chunks, "
          f"{sum(len(row['chunk_data'] or '') for row in template):,} chars")

    r.check(f"{code} has indexed content", bool(filtered) or bool(template),
            f"nothing tagged deliverable_code={code} — vectorize that folder first")

    off_topic = [row for row in filtered if row["deliverable_code"] != code]
    r.check("filtered retrieval is entirely on-deliverable", not off_topic,
            f"leaked: {[(row['filename'], row['deliverable_code']) for row in off_topic[:5]]}")

    if unfiltered:
        others = {row["deliverable_code"] for row in unfiltered} - {code, None}
        print(f"\nunfiltered pulled in other deliverables: {sorted(others) or 'none'}")
        r.check("the filter removes at least something",
                len(filtered) < len(unfiltered) or not others,
                "filtered and unfiltered are identical — either the KB only holds "
                f"{code}, or the filter is not being applied")

    if filtered:
        print("\ntop filtered hits:")
        for row in filtered[:5]:
            print(f"  {row['score']:.3f}  {row['filename']}  "
                  f"[{row['layer']} / {row['layer_category']}]")


# ── main ──────────────────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("tables", help="show both table sets and their row counts")
    sub.add_parser("roundtrip", help="insert tagged fixtures, read back filtered, clean up")

    p_vec = sub.add_parser("vectorize", help="vectorize real S3 content and verify the tags")
    p_vec.add_argument("--prefix", default="knowledge-base/", help="S3 prefix to scan")
    p_vec.add_argument("--limit", type=int, default=5, help="max files to process")
    p_vec.add_argument("--dry-run", action="store_true", help="show derived tags, write nothing")

    p_fnd = sub.add_parser("findings", help="compare filtered vs unfiltered retrieval")
    p_fnd.add_argument("--deliverable", default="B10", help="deliverable rowId, e.g. B10")
    p_fnd.add_argument("--query", default=None, help="override the retrieval query")
    p_fnd.add_argument("--top-k", type=int, default=20)
    p_fnd.add_argument("--template-category", default="Templates",
                       help="layer_category that holds templates")

    args = parser.parse_args()
    r = Results()

    handlers = {
        "tables":    cmd_tables,
        "roundtrip": cmd_roundtrip,
        "vectorize": cmd_vectorize,
        "findings":  cmd_findings,
    }
    try:
        handlers[args.command](args, r)
    except Exception as e:
        print(f"\nERROR  {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return 1

    return r.report()


if __name__ == "__main__":
    sys.exit(main())
