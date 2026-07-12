"""
Run this script once to free up disk space.

Deletes:
  - All recommendations-*.csv files (not needed for coursework)
  - certificates-2012.csv through certificates-2019.csv (old, superseded by newer assessments)

Keeps:
  - certificates-2020.csv through certificates-2026.csv

Run from Windows: python cleanup_data.py
"""

import os
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data" / "raw"

to_delete = (
    list(DATA_DIR.glob("recommendations-*.csv"))
    + [DATA_DIR / f"certificates-{y}.csv" for y in range(2012, 2020)]
)
to_keep = sorted(DATA_DIR.glob("certificates-202*.csv"))

print("Files to DELETE:")
total = 0
for f in sorted(to_delete):
    if f.exists():
        size_gb = f.stat().st_size / 1e9
        total += size_gb
        print(f"  {size_gb:.1f}GB  {f.name}")
    else:
        print(f"  (not found)  {f.name}")

print(f"\nTotal to free: {total:.1f} GB\n")

print("Files to KEEP:")
for f in to_keep:
    if f.exists():
        print(f"  {f.stat().st_size/1e9:.1f}GB  {f.name}")

print()
confirm = input("Type YES to delete, anything else to cancel: ").strip()
if confirm != "YES":
    print("Cancelled.")
else:
    deleted, failed = 0, 0
    for f in to_delete:
        if not f.exists():
            continue
        try:
            f.unlink()
            print(f"deleted: {f.name}")
            deleted += 1
        except Exception as e:
            print(f"FAILED: {f.name} — {e}")
            failed += 1
    print(f"\nDone. Deleted {deleted} files. Failed: {failed}.")
