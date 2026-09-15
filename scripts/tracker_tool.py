#!/usr/bin/env python3
"""Helper for maintaining the "Internship Tracker" sheet in recruiting-tracker.xlsx.

This exists to avoid hand-rolling openpyxl logic (and re-typing large base64
blobs) inside an agent session every time the recruiting-scout task runs.
The agent still owns: locating the file in Drive, downloading/uploading
bytes, and doing the actual job research. This script only owns the local
spreadsheet edit.

Usage:
  tracker_tool.py read <input.xlsx>
      Prints existing "Internship Tracker" rows as JSON, keyed by column
      header. Use this to see what's already logged (dedup by
      "Job Posting Link", or by matching Company + Role Title + Location).

  tracker_tool.py append <input.xlsx> <new_rows.json> <output.xlsx>
      Appends rows from new_rows.json (a JSON list of objects keyed by
      column header) to the Internship Tracker sheet and writes the result
      to <output.xlsx>. Rows are skipped if they duplicate an existing row
      by "Job Posting Link" or by (Company, Role Title, Location). "Date
      Found" is auto-filled with today's date (YYYY-MM-DD) when omitted.
      Prints a JSON summary of what was added vs. skipped, and does not
      touch <input.xlsx>.
"""
import sys
import json
import datetime
import openpyxl

SHEET = "Internship Tracker"
HEADERS = [
    "Date Found", "Company", "Role Title", "Location", "Track", "Category",
    "Deadline to Apply", "Job Posting Link", "Source", "Contact Info", "Notes",
]


def _load(path):
    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET]
    rows = [r for r in ws.iter_rows(min_row=2, values_only=True)
            if any(v not in (None, "") for v in r)]
    return wb, ws, rows


def cmd_read(path):
    _, _, rows = _load(path)
    existing = [dict(zip(HEADERS, r)) for r in rows]
    print(json.dumps(existing, default=str, indent=2))


def cmd_append(path, new_rows_path, out_path):
    wb, ws, rows = _load(path)
    existing_links = {r[7] for r in rows if r[7]}
    existing_triples = {(r[1], r[2], r[3]) for r in rows}

    with open(new_rows_path) as f:
        new_rows = json.load(f)

    today = datetime.date.today().isoformat()
    added, skipped = [], []
    for row in new_rows:
        link = (row.get("Job Posting Link") or "").strip()
        triple = (
            (row.get("Company") or "").strip(),
            (row.get("Role Title") or "").strip(),
            (row.get("Location") or "").strip(),
        )
        if (link and link in existing_links) or triple in existing_triples:
            skipped.append(row)
            continue
        row.setdefault("Date Found", today)
        ws.append([row.get(h, "") for h in HEADERS])
        if link:
            existing_links.add(link)
        existing_triples.add(triple)
        added.append(row)

    wb.save(out_path)
    print(json.dumps(
        {"added": len(added), "skipped": len(skipped),
         "added_rows": added, "skipped_rows": skipped},
        indent=2,
    ))


def main():
    if len(sys.argv) < 3 or sys.argv[1] not in ("read", "append"):
        print(__doc__)
        sys.exit(1)
    if sys.argv[1] == "read":
        cmd_read(sys.argv[2])
    else:
        if len(sys.argv) != 5:
            print(__doc__)
            sys.exit(1)
        cmd_append(sys.argv[2], sys.argv[3], sys.argv[4])


if __name__ == "__main__":
    main()
