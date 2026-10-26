# Integrated Marketing Calendar (Excel) - Implementation Plan

Spec: `marketing_calendar_spec.yaml` (source of truth). Status: DRAFT v0.2.

## What we're building
One `.xlsx` (no macros) with a single data-entry sheet and three formula-driven views.
Based on the "Integrated View" slide: practices as rows, months as columns, each activity
coloured by one of three motions (Create the reason to care / Build momentum / Create opportunity).

| Sheet | What it is | Who touches it |
|---|---|---|
| **Activity Planner** | One row per activity: name, practice, motion, type, start/end, status, owner. Dropdowns + validation. | Data entry (only place) |
| **Annual View** | Practice x month grid (the slide), counts of events/activities per month, per-motion counts, crowding/gap heat-map | Read only |
| **Monthly View** | Pick a month: KPIs, practice x motion matrix, dated activity list | Dropdown only |
| **Practice View** | Pick a practice: 3-lane (motion) timeline across the year, list, KPIs | Dropdown only |
| Lists / Settings / Start Here | Dropdown values, year, rules, instructions | Rare |

Slide principles map to views: Connect -> Practice View, Compound -> Annual grid sequencing, Coordinate -> crowding/gap flags.

## Research result (GitHub + web)
No ready-made equivalent exists. Closest finds:
- **jmcnamara/XlsxWriter** (~4k stars) - chosen build engine (tables, validation, conditional formatting, formulas).
- **trondegil/xlsx-gantt** - Python Gantt-to-xlsx generator; reference only (couldn't read source; not in session scope).
- QuiverTask, gantty - rejected (Electron app / bare Gantt).
- Web tutorials only cover generic calendars; none combine annual + monthly + practice + motion.

Decision: generate the workbook with Python + XlsxWriter from the YAML, so changes (practices, colours, years) are a config edit, not a manual rebuild.

## Phases
1. **Confirm decisions** (open questions below) - 30 min with Dani.
2. **Skeleton**: Lists, Settings, Activity Planner table with validation + sample data from the slide.
3. **Helper columns**: month-overlap flags (M01-M12) and Check column.
4. **Annual View**: slide-style grid, count blocks, heat-map; motion colouring via conditional formatting.
5. **Monthly View** and **Practice View**: selectors + INDEX/SMALL lists (no dynamic arrays).
6. **Polish**: branding, print setup (landscape, 1 page), sheet protection on views, Start Here sheet.
7. **Verify**: LibreOffice recalc (no #REF/#NAME), count reconciliation tests, render screenshots, test in real Excel.
8. **Handoff**: v1.0 file, one-page user guide, how to add a practice.

## Key design choices
- Formulas avoid LET/FILTER/XLOOKUP so it works on locked-down Excel; TEXTJOIN needs Excel 2019+ (fallback defined).
- Everything reads from the table `tblActivities`, so adding rows updates all views automatically.
- Multi-month activities: "span" rule by default (counts in every month it touches).

## Decisions (from Jordan)
- Practices unknown: users add/rename/remove them on the Lists sheet (Excel Table `tblPractices`); dropdowns and all views update automatically. Views are pre-built for 12 practices; blank slots hide.
- Calendar year (Jan-Dec).
- Multi-month activities count in every month they touch.
- Slide palette (navy / blue / teal).
- Several editors: host on SharePoint/OneDrive for co-authoring, lock view sheets (no password), validation + Check column, Change Log sheet.

## Still open
1. Excel version of recipients (assumed Microsoft 365; names-in-cells needs 2019+).
2. Is 12 practices enough headroom?

## Risks
- Per-cell activity names via TEXTJOIN can get long in crowded months -> cap with "+N more" and point to Monthly View.
- Real Excel rendering can differ from LibreOffice -> final check must be in Excel.
