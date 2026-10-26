# Dani Marketing Calendar

Integrated marketing calendar delivered as an Excel workbook (no macros).

- `marketing_calendar.xlsx` - the workbook (Annual / Monthly / Practice views + Activity Planner)
- `marketing_calendar_spec.yaml` - build spec (source of truth)
- `build_workbook.py` - regenerates the workbook: `python3 build_workbook.py marketing_calendar_spec.yaml marketing_calendar.xlsx` (needs `xlsxwriter`, `pyyaml`)
- `IMPLEMENTATION_PLAN.md` - plan, decisions, open questions

Only edit Activity Planner, Lists and Settings in the workbook; the views are formula-driven.
