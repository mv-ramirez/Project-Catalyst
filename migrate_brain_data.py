#!/usr/bin/env python
"""
Backfill the HDI-owned CATALYSTPLATFORM_TB_BRAIN_* tables from the legacy
CATALYSTPLATFORM_BRAIN_* tables.

Why this exists alongside dual-write: BRAIN_WRITE_SET=both only captures ingests
that happen from now on. Everything already vectorized lives only in the legacy
tables, so without a backfill the new tables hold a test-sized sample and the
row counts never line up.

What it does NOT do
-------------------
No DDL. No DROP. It only ever runs INSERT ... SELECT and COUNT(*), so it cannot
destroy anything: if a target row already exists the insert is skipped, and the
legacy tables are never modified or dropped. Dropping them is a separate, manual
decision for later — after the new tables have been verified and lived in
production for a while.

Column handling
---------------
Copies only the columns present in BOTH tables, resolved from SYS.TABLE_COLUMNS
at runtime. That matters because the legacy tables may not have had the tag
columns added yet (POST /db/migrate) — the copy then simply carries what exists
and leaves the new tag columns NULL, rather than failing. Run /db/migrate first
if you want the tags carried across too.

Usage
-----
    # See what would be copied — reads only, writes nothing
    python migrate_brain_data.py plan

    # Copy rows that are not already in the target (safe to re-run)
    python migrate_brain_data.py copy

    # Copy a single table
    python migrate_brain_data.py copy --table vectors

    # Compare row counts between the two sets
    python migrate_brain_data.py verify

Exit code is 0 only if every table verified equal.
"""

import argparse
import sys

import db


# Order matters only for readability; there are no FK constraints between these.
TABLE_KEYS = ["sp_files", "s3_files", "text_chunks", "vectors",
              "files_chunk", "file_registry", "sp_versions"]

# Primary key columns per table, used to skip rows already copied so the run is
# re-runnable without duplicating anything.
PRIMARY_KEYS = {
    "sp_files":      ["item_id"],
    "s3_files":      ["item_id"],
    "text_chunks":   ["item_id", "chunk_id"],
    "vectors":       ["item_id", "chunk_id"],
    "files_chunk":   ["item_id", "chunk_id"],
    "file_registry": ["item_id"],
    "sp_versions":   ["item_id", "version_id"],
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


def _count(cur, table: str):
    try:
        cur.execute(f"SELECT COUNT(*) FROM {db._t(table)}")
        return cur.fetchone()[0]
    except Exception:
        return None


def _plan_table(cur, key: str) -> dict:
    source = db.LEGACY_TABLES[key]
    target = db.HDI_TABLES[key]

    source_cols = _columns_of(cur, source)
    target_cols = _columns_of(cur, target)

    entry = {
        "key":          key,
        "source":       source,
        "target":       target,
        "source_rows":  _count(cur, source),
        "target_rows":  _count(cur, target),
        "shared":       [c for c in source_cols if c in set(target_cols)],
        "source_only":  [c for c in source_cols if c not in set(target_cols)],
        "target_only":  [c for c in target_cols if c not in set(source_cols)],
        "skip":         None,
    }
    if not source_cols:
        entry["skip"] = "legacy table does not exist"
    elif not target_cols:
        entry["skip"] = "HDI table does not exist — deploy the db module first"
    elif not entry["shared"]:
        entry["skip"] = "no columns in common"
    return entry


def cmd_plan(args) -> int:
    conn = db._connect()
    cur  = conn.cursor()
    try:
        keys = [args.table] if args.table else TABLE_KEYS
        print(f"schema: {db._schema}\n")
        blocked = 0
        for key in keys:
            entry = _plan_table(cur, key)
            print(f"{key}")
            print(f"  {entry['source']:<44} {entry['source_rows']} rows")
            print(f"  {entry['target']:<44} {entry['target_rows']} rows")
            if entry["skip"]:
                blocked += 1
                print(f"  SKIP: {entry['skip']}")
            else:
                print(f"  copy {len(entry['shared'])} shared column(s)")
                if entry["source_only"]:
                    print(f"  legacy-only (NOT copied): {', '.join(entry['source_only'])}")
                if entry["target_only"]:
                    print(f"  new columns (left NULL):  {', '.join(entry['target_only'])}")
            print()
        if blocked:
            print(f"{blocked} table(s) cannot be copied — see SKIP above.")
        return 0
    finally:
        cur.close()


def cmd_copy(args) -> int:
    conn = db._connect()
    cur  = conn.cursor()
    keys = [args.table] if args.table else TABLE_KEYS
    failures = 0
    try:
        print(f"schema: {db._schema}\n")
        for key in keys:
            entry = _plan_table(cur, key)
            if entry["skip"]:
                print(f"{key}: SKIP — {entry['skip']}")
                continue

            source, target = entry["source"], entry["target"]
            columns = entry["shared"]
            col_list = ", ".join(f'"{c}"' for c in columns)

            # NOT EXISTS on the primary key makes this idempotent: re-running
            # copies only what is still missing, and never duplicates a row.
            pk = PRIMARY_KEYS[key]
            pk_match = " AND ".join(f't."{c}" = s."{c}"' for c in pk)
            sql = (
                f"INSERT INTO {db._t(target)} ({col_list}) "
                f"SELECT {col_list} FROM {db._t(source)} s "
                f"WHERE NOT EXISTS (SELECT 1 FROM {db._t(target)} t WHERE {pk_match})"
            )

            before = entry["target_rows"]
            try:
                cur.execute(sql)
                conn.commit()
                after = _count(cur, target)
                inserted = (after - before) if (after is not None and before is not None) else "?"
                print(f"{key}: {source} → {target}  +{inserted} rows "
                      f"({before} → {after}; source has {entry['source_rows']})")
            except Exception as e:
                conn.rollback()
                failures += 1
                print(f"{key}: FAILED — {type(e).__name__}: {e}")
        return 1 if failures else 0
    finally:
        cur.close()


def cmd_verify(args) -> int:
    conn = db._connect()
    cur  = conn.cursor()
    keys = [args.table] if args.table else TABLE_KEYS
    mismatches = 0
    try:
        print(f"schema: {db._schema}\n")
        print(f"{'key':<16}{'legacy':>10}{'hdi':>10}   status")
        for key in keys:
            source_rows = _count(cur, db.LEGACY_TABLES[key])
            target_rows = _count(cur, db.HDI_TABLES[key])
            if source_rows is None or target_rows is None:
                status = "missing table"
                mismatches += 1
            elif source_rows == target_rows:
                status = "equal"
            elif target_rows > source_rows:
                # Expected while dual-write is on and new ingests are landing in
                # both — the HDI side can legitimately run ahead of a stale count.
                status = f"hdi ahead by {target_rows - source_rows}"
            else:
                status = f"SHORT by {source_rows - target_rows}"
                mismatches += 1
            print(f"{key:<16}{str(source_rows):>10}{str(target_rows):>10}   {status}")

        print()
        if mismatches:
            print(f"{mismatches} table(s) not yet fully copied. Re-run `copy`.")
            print("Do NOT drop the legacy tables until this reports clean.")
        else:
            print("All tables verified. The legacy tables are still in place and "
                  "remain the rollback path — dropping them is a separate decision.")
        return 1 if mismatches else 0
    finally:
        cur.close()


def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)

    for name, help_text in [
        ("plan",   "show what would be copied (reads only)"),
        ("copy",   "copy missing rows legacy → hdi (idempotent)"),
        ("verify", "compare row counts between the two sets"),
    ]:
        p = sub.add_parser(name, help=help_text)
        p.add_argument("--table", choices=TABLE_KEYS, default=None,
                       help="restrict to one table (default: all)")

    args = parser.parse_args()
    handlers = {"plan": cmd_plan, "copy": cmd_copy, "verify": cmd_verify}
    try:
        return handlers[args.command](args)
    except Exception as e:
        print(f"ERROR  {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
