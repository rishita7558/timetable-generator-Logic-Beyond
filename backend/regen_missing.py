#!/usr/bin/env python3
from app import export_consolidated_semester_timetable, load_all_data

dfs = load_all_data(force_reload=True)

# Regenerate the remaining ones
for (sem, branch) in [(1, "DSAI"), (3, "DSAI")]:
    print(f"\n[REGENERATE] Generating sem{sem}_{branch}...")
    try:
        export_consolidated_semester_timetable(dfs, sem, branch)
        print(f"[OK] Generated sem{sem}_{branch}_timetable.xlsx")
    except Exception as e:
        print(f"[FAIL] {sem}_{branch}: {e}")
