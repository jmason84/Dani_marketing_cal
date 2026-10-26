#!/usr/bin/env python3
"""Build marketing_calendar.xlsx from marketing_calendar_spec.yaml.

Usage: python3 build_workbook.py [spec.yaml] [output.xlsx]
Only the Activity Planner (and Lists / Settings) need data entry; every view is formula-driven.
"""
import calendar
import datetime as dt
import sys

import xlsxwriter
import yaml
from xlsxwriter.utility import xl_col_to_name

SPEC = sys.argv[1] if len(sys.argv) > 1 else "marketing_calendar_spec.yaml"
OUT = sys.argv[2] if len(sys.argv) > 2 else "marketing_calendar.xlsx"
spec = yaml.safe_load(open(SPEC))

YEAR = spec["settings"]["planning_year"]
THRESH = spec["settings"]["collision_threshold"]
MOTIONS = spec["motions"]
PRACTICES = [p["label"] for p in spec["practices"]]
MAXP = spec["max_practices"]
TYPES = spec["activity_types"]
STATUSES = spec["statuses"]
NAVY, BLUE, TEAL = (m["color"] for m in MOTIONS)
MCOL = [m["color"] for m in MOTIONS]
MLABEL = [m["label"] for m in MOTIONS]
MONTHS = [calendar.month_name[i] for i in range(1, 13)]
FILTER_DEFAULT = "All (excl. Cancelled)"
GREY, LIGHT, INK = "#595959", "#F2F2F2", "#1F1F1F"

import os
FIRST = 5
LAST = FIRST + int(os.environ.get("PLANNER_ROWS", 1000)) - 1   # planner data rows (1-based)
NTYPE, NSTAT = 20, 10

PL = "'Activity Planner'!"
wb = xlsxwriter.Workbook(OUT, {"use_future_functions": True})
wb.set_calc_mode("auto")

_fc = {}
def F(**kw):
    kw.setdefault("font_name", "Calibri")
    kw.setdefault("font_size", 10)
    kw.setdefault("valign", "vcenter")
    key = tuple(sorted(kw.items()))
    if key not in _fc:
        _fc[key] = wb.add_format(kw)
    return _fc[key]

def prng(col):
    return f"{PL}${col}${FIRST}:${col}${LAST}"

FLAGS = f"{PL}$O${FIRST}:$Z${LAST}"
def flagcol(m):
    return xl_col_to_name(13 + m)   # m=1 -> O

sheets = {n: wb.add_worksheet(n) for n in [
    "Start Here", "Annual View", "Monthly View", "Practice View",
    "Activity Planner", "Lists", "Settings", "Change Log", "Calc"]}
S, A, M, P, PLN, L, ST, CL, C = (sheets[n] for n in [
    "Start Here", "Annual View", "Monthly View", "Practice View",
    "Activity Planner", "Lists", "Settings", "Change Log", "Calc"])
for ws in sheets.values():
    ws.hide_gridlines(2)
C.hide()

# ---------- defined names ----------
wb.define_name("PlanYear", "=Settings!$B$4")
wb.define_name("StatusFilter", "=Settings!$B$5")
wb.define_name("CollisionThreshold", "=Settings!$B$6")
wb.define_name("PracticeNames", f"=Lists!$B$5:$B${4 + MAXP}")
wb.define_name("MotionNames", "=Lists!$E$5:$E$7")
wb.define_name("MonthList", "=Lists!$M$5:$M$16")
wb.define_name("PracticeList", f"=OFFSET(Lists!$B$5,0,0,MAX(1,COUNTA(Lists!$B$5:$B${4 + MAXP})),1)")
wb.define_name("TypeList", f"=OFFSET(Lists!$G$5,0,0,MAX(1,COUNTA(Lists!$G$5:$G${4 + NTYPE})),1)")
wb.define_name("StatusList", f"=OFFSET(Lists!$I$5,0,0,MAX(1,COUNTA(Lists!$I$5:$I${4 + NSTAT})),1)")
wb.define_name("FilterList", f"=OFFSET(Lists!$K$5,0,0,MAX(1,COUNTIF(Lists!$K$5:$K${6 + NSTAT},\"?*\")),1)")
wb.define_name("SelMonthNum", "='Monthly View'!$K$1")
wb.define_name("SelPractice", "='Practice View'!$B$5")

# common formats
title = F(font_size=22, bold=True, font_color=INK)
kicker = F(font_size=9, bold=True, font_color=TEAL)
sub = F(font_size=11, font_color=GREY)
note = F(font_size=9, italic=True, font_color=GREY)
sect = F(font_size=11, bold=True, font_color=NAVY, bottom=2, bottom_color=NAVY)
hdr = F(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", border=1, border_color="#FFFFFF")
hdrL = F(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="left", indent=1, border=1, border_color="#FFFFFF")
lab = F(bold=True, font_color=INK, indent=1, bottom=1, bottom_color="#D9D9D9")
numc = F(align="center", bottom=1, bottom_color="#D9D9D9")
numcb = F(align="center", bold=True, bottom=1, bottom_color="#D9D9D9")
cell = F(text_wrap=True, valign="top", font_size=8, bold=True, align="center", border=1, border_color="#E7E7E7")
inp = F(bold=True, bg_color="#FFF2CC", border=1, locked=False, align="center")
txt = F(text_wrap=True, bottom=1, bottom_color="#E7E7E7")
datef = F(num_format="dd-mmm-yy", align="center", bottom=1, bottom_color="#E7E7E7")
kpi_l = F(font_size=9, bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", text_wrap=True)
kpi_v = F(font_size=20, bold=True, font_color=NAVY, align="center", bg_color=LIGHT)

def kicker_block(ws, k, t, s, last_col="N"):
    ws.write("A1", k, kicker)
    ws.write("A2", t, title)
    ws.set_row(1, 32)
    ws.write("A3", s, sub)

def motion_cf(ws, rng, first_cell, cols=None):
    """Colour a range by motion label held in first_cell (relative)."""
    for i in range(3):
        ws.conditional_format(rng, {"type": "formula", "criteria": f"={first_cell}=Lists!$E${5 + i}",
                                    "format": wb.add_format({"bg_color": MCOL[i], "font_color": "#FFFFFF", "bold": True})})

# =====================================================================
# LISTS
# =====================================================================
L.set_column("A:A", 3); L.set_column("B:B", 28); L.set_column("C:D", 3)
L.set_column("E:E", 28); L.set_column("F:F", 3); L.set_column("G:G", 26); L.set_column("H:H", 3)
L.set_column("I:I", 18); L.set_column("J:J", 3); L.set_column("K:K", 24); L.set_column("L:L", 3)
L.set_column("M:M", 14)
L.write("B1", "Lists", title)
L.write("B2", "Edit the yellow cells. Add a practice by typing in the next empty yellow cell - every dropdown and view updates.", note)
edit = F(bg_color="#FFF2CC", border=1, border_color="#BFBFBF")
for col, head, n in [("B", "Practices (add / rename)", MAXP), ("G", "Activity types", NTYPE), ("I", "Statuses", NSTAT)]:
    L.write(f"{col}4", head, hdrL)
    vals = {"B": PRACTICES, "G": TYPES, "I": [s["label"] for s in STATUSES]}[col]
    for i in range(n):
        L.write(f"{col}{5 + i}", vals[i] if i < len(vals) else "", edit)
L.write("E4", "Motions (fixed - from the slide)", hdrL)
for i in range(3):
    L.write(f"E{5 + i}", MLABEL[i], F(bold=True, font_color="#FFFFFF", bg_color=MCOL[i], indent=1))
L.write("K4", "Status filter options", hdrL)
L.write("K5", "All (excl. Cancelled)", F(border=1, border_color="#E7E7E7"))
L.write("K6", "All", F(border=1, border_color="#E7E7E7"))
for i in range(NSTAT):
    L.write_formula(f"K{7 + i}", f'=IF($I{5 + i}="","",$I{5 + i})', F(border=1, border_color="#E7E7E7"))
L.write("M4", "Months", hdrL)
for i, mn in enumerate(MONTHS):
    L.write(f"M{5 + i}", mn, F(border=1, border_color="#E7E7E7"))

# =====================================================================
# SETTINGS
# =====================================================================
ST.set_column("A:A", 28); ST.set_column("B:B", 30); ST.set_column("C:C", 70)
ST.write("A1", "Settings", title)
ST.write("A4", "Planning year", lab); ST.write("B4", YEAR, inp)
ST.write("C4", "Calendar year (Jan-Dec). Drives every view.", note)
ST.write("A5", "Status filter", lab); ST.write("B5", FILTER_DEFAULT, inp)
ST.write("C5", "Which activities appear in all views and counts. Default hides Cancelled.", note)
ST.write("A6", "Crowded threshold", lab); ST.write("B6", THRESH, inp)
ST.write("C6", "Annual heat-map turns red when one practice has this many activities in a month.", note)
ST.write("A7", "Multi-month rule", lab); ST.write("B7", "Every month it touches", F(align="center", bg_color=LIGHT))
ST.write("C7", "Fixed: an activity running Jan-Mar counts in Jan, Feb and Mar.", note)
ST.data_validation("B4", {"validate": "integer", "criteria": "between", "minimum": 2020, "maximum": 2100,
                          "error_message": "Enter a 4-digit year."})
ST.data_validation("B5", {"validate": "list", "source": "=FilterList"})
ST.data_validation("B6", {"validate": "integer", "criteria": "between", "minimum": 1, "maximum": 50})

# =====================================================================
# ACTIVITY PLANNER
# =====================================================================
PLN.write("A1", "Activity Planner", title)
PLN.set_row(0, 32)
PLN.write("A2", "One row per activity - this is the only sheet you need to type in. "
                "Leave End date blank for a one-day activity. A red row means the Check column found a problem.", sub)
PLN.write("A3", f'="Rows used: "&COUNTIF($B${FIRST}:$B${LAST},"?*")&" of {LAST - FIRST + 1}   |   Planning year: "&PlanYear&"   |   Showing: "&StatusFilter', note)
heads = ["ID", "Activity", "Practice", "Motion", "Activity type", "Start date", "End date", "Status",
         "Owner", "Audience / channel", "Notes", "Last updated", "Check"]
widths = [9, 34, 20, 26, 18, 12, 12, 13, 16, 22, 34, 13, 24]
for i, w in enumerate(widths):
    PLN.set_column(i, i, w)
PLN.add_table(3, 0, LAST - 1, 12, {
    "name": "tblActivities", "style": "Table Style Light 1",
    "columns": [{"header": h, "header_format": hdr,
                 "format": datef if h in ("Start date", "End date", "Last updated") else F(bottom=1, bottom_color="#E7E7E7")}
                for h in heads]})
PLN.set_row(3, 22)
PLN.set_column("N:AA", 10, None, {"hidden": True})
PLN.write(3, 13, "Include", hdr)
for m in range(1, 13):
    PLN.write(3, 13 + m, f"M{m:02d}", hdr)
PLN.write(3, 26, "Months touched", hdr)
for r in range(FIRST, LAST + 1):
    PLN.write_formula(f"A{r}", f'=IF($B{r}="","","ACT-"&TEXT(ROW()-4,"000"))', F(font_color=GREY, bottom=1, bottom_color="#E7E7E7"))
    PLN.write_formula(f"M{r}", (
        f'=IF(COUNTA($B{r}:$K{r})=0,"",IF($B{r}="","Missing activity name",IF($C{r}="","Missing practice",'
        f'IF(COUNTIF(PracticeNames,$C{r})=0,"Practice not in Lists",IF($D{r}="","Missing motion",'
        f'IF(NOT(ISNUMBER($F{r})),"Missing or invalid start date",IF(AND(ISNUMBER($G{r}),$G{r}<$F{r}),"End date before start",'
        f'IF(AND($N{r}=1,$AA{r}=0),"Outside planning year","OK"))))))))'),
        F(font_size=9, bottom=1, bottom_color="#E7E7E7"))
    PLN.write_formula(f"N{r}", (
        f'=IF(AND($B{r}<>"",$C{r}<>"",$D{r}<>"",ISNUMBER($F{r}),COUNTIF(PracticeNames,$C{r})>0,COUNTIF(MotionNames,$D{r})>0,NOT(AND(ISNUMBER($G{r}),$G{r}<$F{r}))),'
        f'IF(StatusFilter="All (excl. Cancelled)",IF($H{r}="Cancelled",0,1),IF(StatusFilter="All",1,IF($H{r}=StatusFilter,1,0))),0)'))
    for m in range(1, 13):
        PLN.write_formula(f"{flagcol(m)}{r}", (
            f'=IF($N{r}=0,0,IF(AND($F{r}<=EOMONTH(DATE(PlanYear,{m},1),0),'
            f'IF(ISNUMBER($G{r}),$G{r},$F{r})>=DATE(PlanYear,{m},1)),1,0))'))
    PLN.write_formula(f"AA{r}", f"=SUM($O{r}:$Z{r})")
rng = lambda c: f"{c}{FIRST}:{c}{LAST}"
PLN.data_validation(rng("C"), {"validate": "list", "source": "=PracticeList", "error_title": "Practice",
                               "error_message": "Pick a practice from the list (add new ones on the Lists sheet)."})
PLN.data_validation(rng("D"), {"validate": "list", "source": "=MotionNames"})
PLN.data_validation(rng("E"), {"validate": "list", "source": "=TypeList", "error_type": "warning"})
PLN.data_validation(rng("H"), {"validate": "list", "source": "=StatusList"})
PLN.data_validation(rng("F"), {"validate": "date", "criteria": ">=", "value": dt.date(2020, 1, 1),
                               "error_message": "Enter a valid date, e.g. 15-Mar-2027."})
PLN.data_validation(rng("G"), {"validate": "custom", "value": f'=OR($G{FIRST}="",AND(ISNUMBER($G{FIRST}),$G{FIRST}>=$F{FIRST}))',
                               "error_message": "End date must be a date on or after the start date."})
PLN.conditional_format(f"A{FIRST}:M{LAST}", {"type": "formula", "criteria": f'=AND($M{FIRST}<>"",$M{FIRST}<>"OK")',
                                             "format": wb.add_format({"bg_color": "#F8CBAD"})})
motion_cf(PLN, f"D{FIRST}:D{LAST}", f"$D{FIRST}")
for i in range(NSTAT):
    PLN.conditional_format(f"H{FIRST}:H{LAST}", {"type": "formula",
        "criteria": f'=AND($H{FIRST}<>"",$H{FIRST}=Lists!$I${5 + i})',
        "format": wb.add_format({"bg_color": (STATUSES[i]["color"] if i < len(STATUSES) else "#EDEDED")})})
PLN.freeze_panes(4, 2)

# sample data from the slide
sample = spec["sample_data"]
pid = {p["id"]: p["label"] for p in spec["practices"]}
mid = {m["id"]: m["label"] for m in MOTIONS}
for i, s in enumerate(sample):
    r = i + FIRST
    mth = s["month"]
    start = dt.date(YEAR, mth, 5)
    end = dt.date(YEAR, mth, calendar.monthrange(YEAR, mth)[1] - 3)
    if s.get("type") == "Event":
        end = None
        start = dt.date(YEAR, mth, 15)
    if s["activity"] == "Ongoing commentary":
        start, end = dt.date(YEAR, 6, 1), dt.date(YEAR, 8, 31)
    PLN.write(f"B{r}", s["activity"])
    PLN.write(f"C{r}", pid[s["practice"]])
    PLN.write(f"D{r}", mid[s["motion"]])
    PLN.write(f"E{r}", s.get("type", ""))
    PLN.write_datetime(f"F{r}", dt.datetime.combine(start, dt.time()), datef)
    if end:
        PLN.write_datetime(f"G{r}", dt.datetime.combine(end, dt.time()), datef)
    PLN.write(f"H{r}", "Planned")
    PLN.write(f"K{r}", "SAMPLE ROW from the slide - overwrite or delete")

# =====================================================================
# CALC (hidden)
# =====================================================================
C.write("A1", "Calc - hidden helper sheet. Do not edit.", F(bold=True))
BLK = 15
def blk(b):            # first data row (1-based) of block b
    return 3 + b * BLK
blocknames = ["Total", MLABEL[0], MLABEL[1], MLABEL[2], "Events", "Dominant motion (1/2/3)"]
for b, nm in enumerate(blocknames):
    C.write(blk(b) - 1, 0, nm, F(bold=True))
    for m in range(1, 13):
        C.write(blk(b) - 1, m, m)
    for p in range(MAXP):
        r = blk(b) + p
        C.write_formula(f"A{r}", f'=IF(Lists!$B${5 + p}="","",Lists!$B${5 + p})')
        for m in range(1, 13):
            col = xl_col_to_name(m)
            fl = f"{PL}${flagcol(m)}${FIRST}:${flagcol(m)}${LAST}"
            base = f"SUMIFS({fl},{prng('C')},$A{r}"
            if b == 0:
                f = f"=IF($A{r}=\"\",0,{base}))"
            elif b in (1, 2, 3):
                f = f"=IF($A{r}=\"\",0,{base},{prng('D')},Lists!$E${4 + b}))"
            elif b == 4:
                f = f"=IF($A{r}=\"\",0,{base},{prng('E')},\"Event\"))"
            else:
                t, a1, a2, a3 = (f"{col}{blk(k) + p}" for k in range(4))
                f = f"=IF({t}=0,0,IF(AND({a3}>={a2},{a3}>={a1}),3,IF({a2}>={a1},2,1)))"
            C.write_formula(f"{col}{r}", f)
C.write("Q4", "Monthly key", F(bold=True)); C.write("R4", "Practice key", F(bold=True))
for r in range(FIRST, LAST + 1):
    C.write_formula(f"Q{r}", f'=IF(INDEX({PL}$O{r}:$Z{r},SelMonthNum)=1,{PL}$F{r}+ROW()/100000,"")')
    C.write_formula(f"R{r}", f'=IF(AND({PL}$N{r}=1,{PL}$C{r}=SelPractice,{PL}$AA{r}>0),{PL}$F{r}+ROW()/100000,"")')

TOP3 = ("=IF({guard},\"\",TEXTJOIN(CHAR(10),TRUE,IFERROR(INDEX(" + prng("B") +
        ",SMALL(IF({cond},ROW(" + prng("B") + ")-4),{{1,2,3}})),\"\"))&IF({cnt}>3,CHAR(10)&\"+\"&({cnt}-3)&\" more\",\"\"))")
def top3(ws, cell_ref, guard, cond, cnt):
    ws.write_array_formula(f"{cell_ref}:{cell_ref}", "{" + TOP3.format(guard=guard, cond=cond, cnt=cnt) + "}", cell)

# =====================================================================
# ANNUAL VIEW
# =====================================================================
A.set_column("A:A", 24); A.set_column("B:M", 17); A.set_column("N:N", 14)
kicker_block(A, "THE INTEGRATED VIEW", "See the whole motion. Build on what's already happening.",
             "Launches, thought leadership, PR, campaigns, events and client activation in one view, so activity can connect, compound and create sustained opportunity.")
A.write_formula("A4", '="Planning year "&PlanYear&"   |   Showing: "&StatusFilter&"   (change on the Settings sheet)"', note)
A.write("A6", "PRIORITY", hdrL)
for m in range(12):
    A.write(5, 1 + m, MONTHS[m][:3].upper(), hdr)
G0 = 7                                      # first practice row
for p in range(MAXP):
    r = G0 + p
    A.set_row(r - 1, 44)
    A.write_formula(f"A{r}", f'=IF(Lists!$B${5 + p}="","",Lists!$B${5 + p})', F(bold=True, font_size=11, indent=1, text_wrap=True, border=1, border_color="#E7E7E7"))
    for m in range(1, 13):
        col = xl_col_to_name(m)
        fl = f"{PL}${flagcol(m)}${FIRST}:${flagcol(m)}${LAST}"
        top3(A, f"{col}{r}", f'$A{r}=""', f"({prng('C')}=$A{r})*({fl}=1)", f"Calc!{col}{blk(0) + p}")
for m in range(1, 13):
    col = xl_col_to_name(m)
    rg = f"{col}{G0}:{col}{G0 + MAXP - 1}"
    for code, color in ((3, TEAL), (2, BLUE), (1, NAVY)):
        A.conditional_format(rg, {"type": "formula", "criteria": f"=Calc!{col}{blk(5)}={code}",
                                  "format": wb.add_format({"bg_color": color, "font_color": "#FFFFFF"})})
# legend + principles
lr = G0 + MAXP + 1
A.write(f"A{lr}", "LEGEND", F(bold=True, font_size=9, font_color=GREY))
for i in range(3):
    A.merge_range(lr - 1, 1 + i * 2, lr - 1, 2 + i * 2, MLABEL[i], F(bold=True, font_color="#FFFFFF", bg_color=MCOL[i], align="center", font_size=9))
pr = lr + 2
for i, pz in enumerate(spec["principles"]):
    c0 = 1 + i * 4
    A.merge_range(pr - 1, c0, pr - 1, c0 + 3, f"0{i + 1}  {pz['label']}", F(bold=True, font_size=12, font_color=NAVY))
    A.merge_range(pr, c0, pr + 1, c0 + 3, pz["text"], F(text_wrap=True, valign="top", font_size=9, font_color=GREY))
A.write(pr - 1, 0, "WHY THIS VIEW", F(bold=True, font_size=9, font_color=GREY))

# section 1: counts by month
s1 = pr + 4
A.merge_range(s1 - 1, 0, s1 - 1, 13, "ACTIVITIES BY MONTH  (an activity counts in every month it touches; Year column counts each activity once)", sect)
A.write(s1, 0, "", hdrL)
for m in range(12):
    A.write(s1, 1 + m, MONTHS[m][:3].upper(), hdr)
A.write(s1, 13, "YEAR", hdr)
rows1 = [("Total activities", 0, None), ("   of which Events", 4, None), ("   Other activities", None, None),
         (MLABEL[0], 1, 0), (MLABEL[1], 2, 1), (MLABEL[2], 3, 2)]
for i, (name, b, mi) in enumerate(rows1):
    r = s1 + 2 + i
    A.write(f"A{r}", name, lab)
    for m in range(1, 13):
        col = xl_col_to_name(m)
        if b is None:
            A.write_formula(f"{col}{r}", f"={col}{r - 2}-{col}{r - 1}", numc)
        else:
            A.write_formula(f"{col}{r}", f"=SUM(Calc!{col}{blk(b)}:{col}{blk(b) + MAXP - 1})", numcb if i == 0 else numc)
    au = f'{prng("AA")},">0"'
    if i == 0:
        A.write_formula(f"N{r}", f"=COUNTIFS({au})", numcb)
    elif i == 1:
        A.write_formula(f"N{r}", f'=COUNTIFS({au},{prng("E")},"Event")', numc)
    elif i == 2:
        A.write_formula(f"N{r}", f"=N{r - 2}-N{r - 1}", numc)
    else:
        A.write_formula(f"N{r}", f'=COUNTIFS({au},{prng("D")},Lists!$E${5 + mi})', numc)

# section 2: practice x motion
s2 = s1 + 10
A.merge_range(s2 - 1, 0, s2 - 1, 13, "BY PRACTICE AND MOTION  (whole year, each activity counted once)", sect)
for j, h in enumerate(["PRACTICE"] + [m_.upper() for m_ in MLABEL] + ["TOTAL", "EVENTS"]):
    A.write(s2, j, h, hdrL if j == 0 else F(bold=True, font_color="#FFFFFF", bg_color=(MCOL[j - 1] if 1 <= j <= 3 else NAVY), align="center", text_wrap=True, font_size=8))
A.set_row(s2, 30)
for p in range(MAXP):
    r = s2 + 2 + p
    A.write_formula(f"A{r}", f"=$A{G0 + p}", lab)
    au = f'{prng("AA")},">0"'
    for mi in range(3):
        A.write_formula(f"{xl_col_to_name(1 + mi)}{r}", f'=IF($A{r}="","",COUNTIFS({prng("C")},$A{r},{prng("D")},Lists!$E${5 + mi},{au}))', numc)
    A.write_formula(f"E{r}", f'=IF($A{r}="","",COUNTIFS({prng("C")},$A{r},{au}))', numcb)
    A.write_formula(f"F{r}", f'=IF($A{r}="","",COUNTIFS({prng("C")},$A{r},{prng("E")},"Event",{au}))', numc)

# section 3: heat-map
s3 = s2 + MAXP + 4
A.merge_range(s3 - 1, 0, s3 - 1, 13, "CROWDING AND GAPS  (activities per practice per month)", sect)
A.write(s3, 0, "", hdrL)
for m in range(12):
    A.write(s3, 1 + m, MONTHS[m][:3].upper(), hdr)
for p in range(MAXP):
    r = s3 + 2 + p
    A.write_formula(f"A{r}", f"=$A{G0 + p}", lab)
    for m in range(1, 13):
        col = xl_col_to_name(m)
        A.write_formula(f"{col}{r}", f'=IF($A{r}="","",Calc!{col}{blk(0) + p})', F(align="center", border=1, border_color="#FFFFFF"))
hm = f"B{s3 + 2}:M{s3 + 1 + MAXP}"
A.conditional_format(hm, {"type": "formula", "criteria": f'=AND(ISNUMBER(B{s3 + 2}),B{s3 + 2}>=CollisionThreshold)',
                          "format": wb.add_format({"bg_color": "#F4B183", "bold": True})})
A.conditional_format(hm, {"type": "formula", "criteria": f'=AND(ISNUMBER(B{s3 + 2}),B{s3 + 2}=0)',
                          "format": wb.add_format({"bg_color": "#E7E6E6", "font_color": "#A6A6A6"})})
A.conditional_format(hm, {"type": "formula", "criteria": f'=AND(ISNUMBER(B{s3 + 2}),B{s3 + 2}>0)',
                          "format": wb.add_format({"bg_color": "#DDEBF7"})})
A.write_formula(f"A{s3 + 2 + MAXP}", '="Orange = crowded ("&CollisionThreshold&"+ in one month)   Grey = gap   Blue = active"', note)
A.set_landscape(); A.fit_to_pages(1, 0); A.set_paper(9)

# =====================================================================
# MONTHLY VIEW
# =====================================================================
M.set_column("A:A", 28); M.set_column("B:B", 16); M.set_column("C:C", 30); M.set_column("D:D", 22)
M.set_column("E:E", 26); M.set_column("F:F", 16); M.set_column("G:G", 14); M.set_column("H:H", 18)
M.set_column("K:K", 6, None, {"hidden": True})
kicker_block(M, "MONTHLY VIEW", "One month, in detail.", "Pick a month to see what is happening, for which practice, and in which motion.")
M.write("A5", "Select month:", F(bold=True, align="right"))
M.write("B5", "March", inp)
M.data_validation("B5", {"validate": "list", "source": "=MonthList"})
M.write_formula("K1", "=MATCH($B$5,MonthList,0)")
M.write_formula("C5", '="Planning year "&PlanYear&"   |   Showing: "&StatusFilter', note)
for i, h in enumerate(["TOTAL ACTIVITIES", "EVENTS", MLABEL[0].upper(), MLABEL[1].upper(), MLABEL[2].upper()]):
    M.write(6, i, h, kpi_l if i < 2 else F(font_size=9, bold=True, font_color="#FFFFFF", bg_color=MCOL[i - 2], align="center", text_wrap=True))
M.set_row(6, 28); M.set_row(7, 36)
fm = f"INDEX({FLAGS},0,SelMonthNum)"
M.write_formula("A8", f"=SUMPRODUCT({fm})", kpi_v)
M.write_formula("B8", f'=SUMPRODUCT({fm}*({prng("E")}="Event"))', kpi_v)
for mi in range(3):
    M.write_formula(f"{xl_col_to_name(2 + mi)}8", f'=SUMPRODUCT({fm}*({prng("D")}=Lists!$E${5 + mi}))', kpi_v)
M.merge_range("A11:H11", "PRACTICE BY MOTION", sect)
for j, h in enumerate(["PRACTICE"] + MLABEL + ["TOTAL"]):
    M.write(11, j, h, hdrL if j == 0 else F(bold=True, font_color="#FFFFFF", bg_color=(MCOL[j - 1] if 1 <= j <= 3 else NAVY), align="center", text_wrap=True, font_size=9))
M.set_row(11, 30)
for p in range(MAXP):
    r = 13 + p
    M.write_formula(f"A{r}", f'=IF(Lists!$B${5 + p}="","",Lists!$B${5 + p})', lab)
    for mi in range(3):
        M.write_formula(f"{xl_col_to_name(1 + mi)}{r}", f'=IF($A{r}="","",SUMPRODUCT({fm}*({prng("C")}=$A{r})*({prng("D")}=Lists!$E${5 + mi})))', numc)
    M.write_formula(f"E{r}", f'=IF($A{r}="","",SUM(B{r}:D{r}))', numcb)
wk = 13 + MAXP + 1                      # week section header row (1-based)
M.merge_range(wk - 1, 0, wk - 1, 7, "WEEK BY WEEK", sect)
M.write(wk, 0, "", hdrL)
for w in range(5):
    start = f"DATE(PlanYear,SelMonthNum,{1 + 7 * w})"
    M.write_formula(wk, 1 + w, f'=IF({1 + 7 * w}>DAY(EOMONTH(DATE(PlanYear,SelMonthNum,1),0)),"","Week {w + 1}  ("&{1 + 7 * w}&"-"&MIN({7 * w + 7},DAY(EOMONTH(DATE(PlanYear,SelMonthNum,1),0)))&")")', hdr)
M.write(wk + 1, 0, "Active that week", lab)
M.write(wk + 2, 0, "What", lab)
M.set_row(wk + 2, 96)
endeff = f"IF(ISNUMBER({prng('G')}),{prng('G')},{prng('F')})"
for w in range(5):
    col = xl_col_to_name(1 + w)
    ws_, we_ = f"DATE(PlanYear,SelMonthNum,{1 + 7 * w})", f"MIN(DATE(PlanYear,SelMonthNum,{7 * w + 7}),EOMONTH(DATE(PlanYear,SelMonthNum,1),0))"
    cond = f"({prng('N')}=1)*({prng('F')}<={we_})*({endeff}>={ws_})"
    M.write_formula(f"{col}{wk + 2}", f'=IF({1 + 7 * w}>DAY(EOMONTH(DATE(PlanYear,SelMonthNum,1),0)),"",SUMPRODUCT({cond}))', numcb)
    TOP5 = ("=IF({g}>DAY(EOMONTH(DATE(PlanYear,SelMonthNum,1),0)),\"\",TEXTJOIN(CHAR(10),TRUE,IFERROR(INDEX(" + prng("B") +
            ",SMALL(IF({c},ROW(" + prng("B") + ")-4),{{1,2,3,4,5}})),\"\"))&IF({col}{r}>5,CHAR(10)&\"+\"&({col}{r}-5)&\" more\",\"\"))")
    M.write_array_formula(f"{col}{wk + 3}:{col}{wk + 3}", "{" + TOP5.format(g=1 + 7 * w, c=cond, col=col, r=wk + 2) + "}", cell)
li = wk + 5                              # list section header row
M.merge_range(li - 1, 0, li - 1, 7, "", sect)
M.write_formula(li - 1, 0, '="ACTIVITIES IN "&UPPER($B$5)&" (sorted by start date)"', sect)
for j, h in enumerate(["ACTIVITY", "START", "END", "PRACTICE", "MOTION", "TYPE", "STATUS", "OWNER"]):
    M.write(li, j, h, hdrL if j == 0 else hdr)
NL = 40
for k in range(1, NL + 1):
    r = li + 1 + k
    M.write_formula(f"K{r}", f'=IFERROR(MATCH(SMALL(Calc!$Q${FIRST}:$Q${LAST},{k}),Calc!$Q${FIRST}:$Q${LAST},0),"")')
    for j, (col, kind) in enumerate([("B", "t"), ("F", "d"), ("G", "d"), ("C", "t"), ("D", "t"), ("E", "t"), ("H", "t"), ("I", "t")]):
        ix = f"INDEX({prng(col)},$K{r})"
        f = f'=IF($K{r}="","",IF({ix}="","",{ix}))'
        M.write_formula(r - 1, j, f, datef if kind == "d" else txt)
M.write_formula(li + NL + 2, 0, f'=IF(COUNT(Calc!$Q${FIRST}:$Q${LAST})>{NL},"Showing the first {NL} of "&COUNT(Calc!$Q${FIRST}:$Q${LAST})&" activities - filter the Activity Planner for the rest.","")', note)
motion_cf(M, f"E{li + 2}:E{li + 1 + NL}", f"$E{li + 2}")
M.set_landscape(); M.fit_to_pages(1, 0)

# =====================================================================
# PRACTICE VIEW
# =====================================================================
P.set_column("A:A", 30); P.set_column("B:M", 17); P.set_column("O:O", 6, None, {"hidden": True})
kicker_block(P, "PRACTICE VIEW", "One practice, across the year.", "Pick a practice to see its three motions month by month, and every activity behind them.")
P.write("A5", "Select practice:", F(bold=True, align="right"))
P.write("B5", PRACTICES[0], inp)
P.data_validation("B5", {"validate": "list", "source": "=PracticeList"})
P.write_formula("C5", '="Planning year "&PlanYear&"   |   Showing: "&StatusFilter', note)
kp = ["TOTAL ACTIVITIES", "EVENTS", MLABEL[0].upper(), MLABEL[1].upper(), MLABEL[2].upper(), "BUSIEST MONTH"]
for i, h in enumerate(kp):
    P.write(6, i, h, kpi_l if i in (0, 1, 5) else F(font_size=9, bold=True, font_color="#FFFFFF", bg_color=MCOL[i - 2], align="center", text_wrap=True))
P.set_row(6, 28); P.set_row(7, 36)
PK = f"Calc!$R${FIRST}:$R${LAST}"
P.write_formula("A8", f"=COUNT({PK})", kpi_v)
P.write_formula("B8", f'=SUMPRODUCT(--ISNUMBER({PK}),--({prng("E")}="Event"))', kpi_v)
for mi in range(3):
    P.write_formula(f"{xl_col_to_name(2 + mi)}8", f'=SUMPRODUCT(--ISNUMBER({PK}),--({prng("D")}=Lists!$E${5 + mi}))', kpi_v)
P.write_formula("F8", '=IF(MAX($B$17:$M$17)=0,"-",INDEX(MonthList,MATCH(MAX($B$17:$M$17),$B$17:$M$17,0)))', F(font_size=14, bold=True, font_color=NAVY, align="center", bg_color=LIGHT))
P.merge_range("A10:M10", "THREE MOTIONS ACROSS THE YEAR", sect)
P.write(10, 0, "MOTION", hdrL)
for m in range(12):
    P.write(10, 1 + m, MONTHS[m][:3].upper(), hdr)
for i in range(3):
    r = 12 + i
    P.set_row(r - 1, 58)
    P.write_formula(f"A{r}", f"=Lists!$E${5 + i}", F(bold=True, font_color="#FFFFFF", bg_color=MCOL[i], indent=1, text_wrap=True))
    for m in range(1, 13):
        col = xl_col_to_name(m)
        fl = f"{PL}${flagcol(m)}${FIRST}:${flagcol(m)}${LAST}"
        cnt = f'SUMIFS({fl},{prng("C")},SelPractice,{prng("D")},$A{r})'
        top3(P, f"{col}{r}", 'SelPractice=""', f"({prng('C')}=SelPractice)*({prng('D')}=$A{r})*({fl}=1)", cnt)
    P.conditional_format(f"B{r}:M{r}", {"type": "formula", "criteria": f"=LEN(B{r})>0",
                                        "format": wb.add_format({"bg_color": MCOL[i], "font_color": "#FFFFFF"})})
P.write("A16", "", F())
P.write("A17", "Activities per month", lab)
for m in range(1, 13):
    col = xl_col_to_name(m)
    fl = f"{PL}${flagcol(m)}${FIRST}:${flagcol(m)}${LAST}"
    P.write_formula(f"{col}17", f'=SUMIFS({fl},{prng("C")},SelPractice)', numcb)
P.conditional_format("B17:M17", {"type": "data_bar", "bar_color": BLUE, "bar_solid": True})
P.merge_range("A19:M19", "ACTIVITY LIST (sorted by start date)", sect)
for j, h in enumerate(["ACTIVITY", "START", "END", "MOTION", "TYPE", "STATUS", "OWNER", "AUDIENCE / CHANNEL"]):
    P.write(19, j, h, hdrL if j == 0 else hdr)
NLP = 60
for k in range(1, NLP + 1):
    r = 20 + k
    P.write_formula(f"O{r}", f'=IFERROR(MATCH(SMALL({PK},{k}),{PK},0),"")')
    for j, (col, kind) in enumerate([("B", "t"), ("F", "d"), ("G", "d"), ("D", "t"), ("E", "t"), ("H", "t"), ("I", "t"), ("J", "t")]):
        ix = f"INDEX({prng(col)},$O{r})"
        P.write_formula(r - 1, j, f'=IF($O{r}="","",IF({ix}="","",{ix}))', datef if kind == "d" else txt)
P.write_formula(20 + NLP + 1, 0, f'=IF(COUNT({PK})>{NLP},"Showing the first {NLP} of "&COUNT({PK})&" activities.","")', note)
motion_cf(P, f"D21:D{20 + NLP}", "$D21")
P.set_landscape(); P.fit_to_pages(1, 0)

# =====================================================================
# CHANGE LOG
# =====================================================================
CL.set_column("A:A", 14); CL.set_column("B:B", 22); CL.set_column("C:C", 90)
CL.write("A1", "Change Log", title)
CL.write("A2", "Several people edit this file - note structural changes here (new practice, renamed type, year rolled forward).", note)
for j, h in enumerate(["Date", "Who", "What changed"]):
    CL.write(3, j, h, hdrL)
CL.write_datetime(4, 0, dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0), datef)
CL.write(4, 1, "Workbook generator"); CL.write(4, 2, "v0.2 created from marketing_calendar_spec.yaml")

# =====================================================================
# START HERE
# =====================================================================
S.set_column("A:A", 3); S.set_column("B:B", 30); S.set_column("C:C", 100)
S.write("B1", "THE INTEGRATED VIEW", kicker)
S.write("B2", "Integrated Marketing Calendar", title)
S.set_row(1, 32)
S.write("B3", "One place to plan the whole motion - connect it, compound it, coordinate it.", sub)
rows = [
    ("How it works", ""),
    ("1. Enter activities", "Go to Activity Planner. One row per activity: name, practice, motion, type, start (and end) date, status. Everything else updates itself."),
    ("2. Read the views", "Annual View = the whole year by practice and month. Monthly View = pick a month. Practice View = pick a practice."),
    ("3. Add a practice", "Lists sheet > Practices: type the new name in the next yellow cell. It appears in the dropdowns and all views (room for 12)."),
    ("4. Roll the year", "Settings > Planning year. Update dates in the planner; clear sample rows first."),
    ("", ""),
    ("The three motions", ""),
    (MLABEL[0], "Why should the market care? Point of view, research, market insight."),
    (MLABEL[1], "Keep the story moving. PR, campaigns, launches, events, executive visibility."),
    (MLABEL[2], "Turn attention into conversations. Client activation, webinars, roundtables, outreach, follow-up."),
    ("", ""),
    ("The three principles", ""),
]
for pz in spec["principles"]:
    rows.append((pz["label"], pz["text"]))
rows += [("", ""), ("Good to know", ""),
         ("Multi-month activities", "An activity from 1-Mar to 15-May counts in March, April and May. Year totals count it once."),
         ("Status filter", "Settings > Status filter. Default hides Cancelled. Applies to every view."),
         ("Red row in planner", "The Check column explains what is missing (practice, motion, date...)."),
         ("Several editors", "Keep the file on SharePoint/OneDrive so people can edit together. Only edit Activity Planner, Lists and Settings. View sheets are locked (no password)."),
         ("Sample data", "The 19 sample rows are from the original slide. Delete or overwrite them. Do not delete rows - clear their contents."),
         ("Capacity", f"{LAST - FIRST + 1} activity rows, 12 practices, 20 activity types.")]
r = 4
for a, b in rows:
    if b == "" and a:
        S.write(r, 1, a.upper(), F(bold=True, font_size=11, font_color=NAVY, bottom=2, bottom_color=NAVY)); S.write(r, 2, "", F(bottom=2, bottom_color=NAVY))
    elif a:
        fmt = F(bold=True, indent=1)
        for i in range(3):
            if a == MLABEL[i]:
                fmt = F(bold=True, font_color="#FFFFFF", bg_color=MCOL[i], indent=1)
        S.write(r, 1, a, fmt)
        S.write(r, 2, b, F(text_wrap=True))
    r += 1

# protection: views are locked (no password); inputs are unlocked
for ws in (S, A, M, P, C):
    ws.protect("", {"select_locked_cells": True, "select_unlocked_cells": True, "format_columns": True, "format_rows": True})
for ws in (S, PLN, L, ST, CL):
    ws.set_landscape(); ws.fit_to_pages(1, 0); ws.set_paper(9)
PLN.repeat_rows(3)
S.activate()
wb.close()
print("wrote", OUT)
