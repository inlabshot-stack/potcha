#!/usr/bin/env python3
"""
Potcha Ops — Business Formulas Workbook builder.
Extracts verbatim formulas from the app source and re-implements them live in Excel.
Output: /home/z/my-project/Potcha-Ops-Business-Formulas.xlsx (workspace root, not public/).
"""
import sys, os
XLSX_SKILL_DIR = "/home/z/my-project/skills/xlsx"
for sub in [XLSX_SKILL_DIR, os.path.join(XLSX_SKILL_DIR, "templates")]:
    if sub not in sys.path:
        sys.path.insert(0, sub)

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from templates.base import (
    FONT_NAME, NEUTRAL_900, NEUTRAL_600, NEUTRAL_200,
    ACCENT_POSITIVE, ACCENT_NEGATIVE, ACCENT_WARNING,
)

HEADER_FILL = PatternFill("solid", fgColor=NEUTRAL_900)
INPUT_FILL = PatternFill("solid", fgColor="FFF8E1")
CALC_FILL = PatternFill("solid", fgColor="F2F4F7")
TITLE_FONT = Font(name=FONT_NAME, size=16, bold=True, color=NEUTRAL_900)
H_FONT = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
B_FONT = Font(name=FONT_NAME, size=10, color=NEUTRAL_900)
MUTED = Font(name=FONT_NAME, size=9, italic=True, color=NEUTRAL_600)
CODE_FONT = Font(name="Consolas", size=9, color=NEUTRAL_900)
THIN = Side(style="thin", color=NEUTRAL_200)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(vertical="top", wrap_text=True)
CTR = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
DATEFMT = "yyyy-mm-dd"; DATETIMEFMT = "yyyy-mm-dd hh:mm"
NUMFMT = "#,##0"; PCTFMT = "0%"

def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL; cell.font = H_FONT; cell.alignment = CTR; cell.border = BORDER

def title_block(ws, title, subtitle, span=6):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    c = ws.cell(row=1, column=1, value=title); c.font = TITLE_FONT; c.alignment = LEFT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
    c = ws.cell(row=2, column=1, value=subtitle); c.font = MUTED; c.alignment = LEFT
    ws.row_dimensions[1].height = 22

def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w

red_fill = PatternFill("solid", fgColor="FDEDEC")
red_font = Font(name=FONT_NAME, size=10, bold=True, color=ACCENT_NEGATIVE)
amber_fill = PatternFill("solid", fgColor="FEF9E7")
green_fill = PatternFill("solid", fgColor="E8F5E9")

wb = Workbook(); wb.properties.creator = "Z.ai"
wb.properties.title = "Potcha Ops — Business Formulas"
wb.remove(wb.active)

# ============ SHEET 1: Overview ============
ws = wb.create_sheet("Overview")
title_block(ws, "Potcha Ops — Business Formulas",
            "Every formula the app uses around task delays, critical issues, staff workload, readiness & finance — extracted verbatim from source, re-implemented live.", span=6)
widths(ws, {"A": 3, "B": 28, "C": 72, "D": 22, "E": 4, "F": 4})
intro = [
    ("What this is", "A working Excel reference for every business rule Potcha Ops computes server-side. Each live sheet lets you change the inputs (amber cells) and watch the result update. The Formula Reference sheet lists the exact source line for each."),
    ("Where it came from", "Formulas are copied from the actual codebase — not invented. Source files: src/lib/readiness.ts, src/lib/finance.ts, src/lib/constants.ts, src/app/api/dashboard/route.ts, src/app/api/tasks/route.ts."),
]
r = 4
for lbl, txt in intro:
    ws.cell(row=r, column=2, value=lbl).font = Font(name=FONT_NAME, size=10, bold=True, color=NEUTRAL_900)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    c = ws.cell(row=r, column=3, value=txt); c.font = B_FONT; c.alignment = WRAP
    ws.row_dimensions[r].height = 30
    r += 1
r += 1
ws.cell(row=r, column=2, value="Sheet").font = H_FONT; ws.cell(row=r, column=2).fill = HEADER_FILL; ws.cell(row=r, column=2).border = BORDER
ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
ws.cell(row=r, column=3, value="Purpose").font = H_FONT; ws.cell(row=r, column=3).fill = HEADER_FILL; ws.cell(row=r, column=3).border = BORDER
ws.cell(row=r, column=6, value="Live formulas").font = H_FONT; ws.cell(row=r, column=6).fill = HEADER_FILL; ws.cell(row=r, column=6).border = BORDER; ws.cell(row=r, column=6).alignment = CTR
r += 1
sheet_index = [
    ("Task Delay & Overdue", "When a task counts as overdue / due-today / upcoming, days overdue, and the payment-overdue rule.", "6"),
    ("Critical Issues", "Open critical issue detection, open high/critical, and the composite event-at-risk rule.", "4"),
    ("Staff & Workload", "Per-staff task counts (assigned, overdue, critical, completed) and completion %, plus run-sheet completion.", "6"),
    ("Readiness Score", "The full 6-criterion weighted model (100 pts) as a live calculator with band classification.", "8"),
    ("Finance Formulas", "Collected, outstanding, expenses, profit & margin — all live from a payments table.", "9"),
    ("Formula Reference", "Master table: every formula, plain-English rule, exact code expression, source file:line.", "—"),
]
for name, purpose, n in sheet_index:
    ws.cell(row=r, column=2, value=name).font = Font(name=FONT_NAME, size=10, bold=True, color=NEUTRAL_900); ws.cell(row=r, column=2).border = BORDER
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    c = ws.cell(row=r, column=3, value=purpose); c.font = B_FONT; c.alignment = WRAP; c.border = BORDER
    ws.cell(row=r, column=6, value=n).font = MUTED if n == "—" else B_FONT; ws.cell(row=r, column=6).alignment = CTR; ws.cell(row=r, column=6).border = BORDER
    ws.row_dimensions[r].height = 28
    r += 1
r += 1
ws.cell(row=r, column=2, value="Legend").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900); r += 1
legend = [
    ("Amber cell", "Editable input — change me", INPUT_FILL),
    ("Gray cell", "Computed by Excel formula — do not type", CALC_FILL),
    ("Red", "Overdue / critical / at-risk", red_fill),
    ("Amber", "Attention / due soon", amber_fill),
    ("Green", "Done / on track / resolved", green_fill),
]
for lbl, desc, fill in legend:
    c1 = ws.cell(row=r, column=2, value=lbl); c1.fill = fill; c1.font = B_FONT; c1.border = BORDER
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    c2 = ws.cell(row=r, column=3, value=desc); c2.font = B_FONT; c2.alignment = WRAP; c2.border = BORDER
    r += 1
ws.freeze_panes = "A4"

# ============ SHEET 2: Task Delay & Overdue ============
ws = wb.create_sheet("Task Delay & Overdue")
title_block(ws, "Task Delay & Overdue",
            "Source: api/tasks/route.ts:60 (isOverdue flag), api/dashboard/route.ts:126-131 (overdue + due-today), lib/finance.ts:30 (payment overdue).", span=8)
widths(ws, {"A": 4, "B": 32, "C": 14, "D": 16, "E": 20, "F": 22, "G": 14, "H": 14})
ws["B4"] = "now (with time)"; ws["C4"] = "=NOW()"; ws["C4"].number_format = DATETIMEFMT; ws["C4"].fill = INPUT_FILL; ws["C4"].border = BORDER; ws["C4"].alignment = CTR
ws["B5"] = "start of today"; ws["C5"] = "=TODAY()"; ws["C5"].number_format = DATEFMT; ws["C5"].fill = INPUT_FILL; ws["C5"].border = BORDER; ws["C5"].alignment = CTR
ws["B6"] = "end of today (exclusive)"; ws["C6"] = "=TODAY()+1"; ws["C6"].number_format = DATEFMT; ws["C6"].fill = CALC_FILL; ws["C6"].border = BORDER; ws["C6"].alignment = CTR
for rr in (4, 5, 6):
    ws.cell(row=rr, column=2).font = Font(name=FONT_NAME, size=10, bold=True, color=NEUTRAL_900); ws.cell(row=rr, column=2).alignment = RIGHT
HR = 9
headers = ["#", "Task", "Due date", "Status", "Overdue? (flag, uses now)", "Overdue? (dashboard, startOfToday)", "Days overdue", "Due today?"]
for i, h in enumerate(headers): ws.cell(row=HR, column=i+1, value=h)
style_header(ws, HR, 8)
sample = [
    ("Confirm floral arch design",       "=TODAY()-5", "TODO"),
    ("Pay photographer deposit",         "=TODAY()-1", "IN_PROGRESS"),
    ("Send final guest count to caterer","=TODAY()",   "TODO"),
    ("Order ceremony signage",           "=TODAY()+2", "TODO"),
    ("Confirm transport for bridal party","=TODAY()+7","TODO"),
    ("Lock sound-check schedule",        "=TODAY()-3", "COMPLETED"),
    ("Final dress fitting",              "=TODAY()-2", "BLOCKED"),
    ("Brief ushers",                     "=TODAY()+1", "TODO"),
]
r = HR + 1
for i, (name, due, status) in enumerate(sample, 1):
    ws.cell(row=r, column=1, value=i).alignment = CTR
    ws.cell(row=r, column=2, value=name).font = B_FONT
    c = ws.cell(row=r, column=3, value=due); c.number_format = DATEFMT; c.alignment = CTR
    ws.cell(row=r, column=4, value=status).alignment = CTR
    ws.cell(row=r, column=5, value=f'=IF(AND($C{r}<>"",$C{r}<$C$4,$D{r}<>"COMPLETED"),"YES","NO")').alignment = CTR
    ws.cell(row=r, column=6, value=f'=IF(AND($D{r}<>"COMPLETED",$C{r}<>"",$C{r}<$C$5),"YES","NO")').alignment = CTR
    ws.cell(row=r, column=7, value=f'=IF(AND($D{r}<>"COMPLETED",$C{r}<$C$5),$C$5-$C{r},0)').alignment = CTR
    ws.cell(row=r, column=8, value=f'=IF(AND($D{r}<>"COMPLETED",$C{r}>=$C$5,$C{r}<$C$6),"YES","NO")').alignment = CTR
    for col in range(1, 9):
        ws.cell(row=r, column=col).border = BORDER
        if col >= 5: ws.cell(row=r, column=col).fill = CALC_FILL; ws.cell(row=r, column=col).font = B_FONT
    r += 1
last = r - 1
for col in ["E", "F", "H"]:
    ws.conditional_formatting.add(f"{col}{HR+1}:{col}{last}", CellIsRule(operator="equal", formula=['"YES"'], fill=red_fill, font=red_font))
sr = last + 2
ws.cell(row=sr, column=2, value="Summary (live)").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
kpis = [
    ("Overdue tasks (dashboard rule)", f'=COUNTIF(F{HR+1}:F{last},"YES")', True),
    ("Overdue tasks (display flag)",   f'=COUNTIF(E{HR+1}:E{last},"YES")', True),
    ("Due today",                       f'=COUNTIF(H{HR+1}:H{last},"YES")', False),
    ("Completed (excluded from overdue)", f'=COUNTIF(D{HR+1}:D{last},"COMPLETED")', False),
]
for i, (lbl, fml, red) in enumerate(kpis):
    rr = sr + 1 + i
    ws.cell(row=rr, column=2, value=lbl).font = B_FONT
    c = ws.cell(row=rr, column=3, value=fml); c.fill = CALC_FILL; c.border = BORDER; c.font = Font(name=FONT_NAME, size=11, bold=True, color=ACCENT_NEGATIVE if red else NEUTRAL_900); c.alignment = CTR
cr = sr + len(kpis) + 3
ws.cell(row=cr, column=2, value="Exact source expressions").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
refs = [
    ("Task isOverdue flag",     "t.dueDate !== null && t.dueDate < now && t.status !== 'COMPLETED'", "api/tasks/route.ts:60"),
    ("Dashboard overdue task",  "t.status !== 'COMPLETED' && t.dueDate && t.dueDate < startOfToday",  "api/dashboard/route.ts:126-128"),
    ("Task due today",          "t.status !== 'COMPLETED' && t.dueDate >= startOfToday && t.dueDate <= endOfToday", "api/dashboard/route.ts:129-131"),
    ("Overdue query filter",    "where.dueDate = { lt: startOfToday }",                              "api/tasks/route.ts:40"),
    ("Payment overdue",         "p.status === 'PENDING' && p.dueDate !== null && p.dueDate < now",   "lib/finance.ts:30"),
    ("Payment due <= 14 days",  "p.dueDate >= startOfToday && p.dueDate <= now + 14*86400000",       "api/dashboard/route.ts:121-122"),
]
ws.cell(row=cr+1, column=2, value="Rule").font = H_FONT; ws.cell(row=cr+1, column=2).fill = HEADER_FILL
ws.merge_cells(start_row=cr+1, start_column=3, end_row=cr+1, end_column=6)
ws.cell(row=cr+1, column=3, value="Code expression").font = H_FONT; ws.cell(row=cr+1, column=3).fill = HEADER_FILL
ws.merge_cells(start_row=cr+1, start_column=7, end_row=cr+1, end_column=8)
ws.cell(row=cr+1, column=7, value="Source").font = H_FONT; ws.cell(row=cr+1, column=7).fill = HEADER_FILL
for i, (rule, expr, src) in enumerate(refs):
    rr = cr + 2 + i
    ws.cell(row=rr, column=2, value=rule).font = B_FONT; ws.cell(row=rr, column=2).border = BORDER
    ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=6)
    c = ws.cell(row=rr, column=3, value=expr); c.font = CODE_FONT; c.alignment = WRAP; c.border = BORDER
    ws.merge_cells(start_row=rr, start_column=7, end_row=rr, end_column=8)
    ws.cell(row=rr, column=7, value=src).font = MUTED; ws.cell(row=rr, column=7).border = BORDER
ws.freeze_panes = "B10"

# ============ SHEET 3: Critical Issues ============
ws = wb.create_sheet("Critical Issues")
title_block(ws, "Critical Issues",
            "Source: api/dashboard/route.ts:132-134 (open critical), lib/readiness.ts:151-153 (open high/critical), api/dashboard/route.ts:102-109 (event at risk).", span=8)
widths(ws, {"A": 4, "B": 30, "C": 14, "D": 16, "E": 18, "F": 20, "G": 14, "H": 14})
ws["B4"] = "now"; ws["C4"] = "=NOW()"; ws["C4"].number_format = DATETIMEFMT; ws["C4"].fill = INPUT_FILL; ws["C4"].border = BORDER; ws["C4"].alignment = CTR
ws["B5"] = "at-risk window (days)"; ws["C5"] = 30; ws["C5"].fill = INPUT_FILL; ws["C5"].border = BORDER; ws["C5"].alignment = CTR
ws.cell(row=4, column=2).font = Font(name=FONT_NAME, size=10, bold=True); ws.cell(row=4, column=2).alignment = RIGHT
ws.cell(row=5, column=2).font = Font(name=FONT_NAME, size=10, bold=True); ws.cell(row=5, column=2).alignment = RIGHT
HR = 8
ihdrs = ["#", "Issue", "Priority", "Status", "Open critical?", "Open high/critical?", "Days open", "Created"]
for i, h in enumerate(ihdrs): ws.cell(row=HR, column=i+1, value=h)
style_header(ws, HR, 8)
issues = [
    ("Venue double-booked rumor",      "CRITICAL", "OPEN",        "=TODAY()-4"),
    ("Florist deposit not received",   "HIGH",     "OPEN",        "=TODAY()-2"),
    ("Cake design mismatch",           "MEDIUM",   "IN_PROGRESS", "=TODAY()-6"),
    ("Sound system failure on test",   "CRITICAL", "IN_PROGRESS", "=TODAY()-1"),
    ("Late guest RSVPs",               "LOW",      "OPEN",        "=TODAY()-9"),
    ("Generator backup confirmed",     "HIGH",     "RESOLVED",    "=TODAY()-12"),
    ("Transport route permit",         "CRITICAL", "RESOLVED",    "=TODAY()-8"),
]
r = HR + 1
for i, (name, pri, status, created) in enumerate(issues, 1):
    ws.cell(row=r, column=1, value=i).alignment = CTR
    ws.cell(row=r, column=2, value=name).font = B_FONT
    ws.cell(row=r, column=3, value=pri).alignment = CTR
    ws.cell(row=r, column=4, value=status).alignment = CTR
    c = ws.cell(row=r, column=8, value=created); c.number_format = DATEFMT; c.alignment = CTR
    ws.cell(row=r, column=5, value=f'=IF(AND($C{r}="CRITICAL",$D{r}<>"RESOLVED"),"YES","NO")').alignment = CTR
    ws.cell(row=r, column=6, value=f'=IF(AND(OR($C{r}="HIGH",$C{r}="CRITICAL"),$D{r}<>"RESOLVED"),"YES","NO")').alignment = CTR
    ws.cell(row=r, column=7, value=f'=IF($D{r}<>"RESOLVED",$C$4-$H{r},"")').alignment = CTR
    for col in range(1, 9):
        ws.cell(row=r, column=col).border = BORDER
        if col in (5,6,7): ws.cell(row=r, column=col).fill = CALC_FILL; ws.cell(row=r, column=col).font = B_FONT
    r += 1
ilast = r - 1
for col in ["E", "F"]:
    ws.conditional_formatting.add(f"{col}{HR+1}:{col}{ilast}", CellIsRule(operator="equal", formula=['"YES"'], fill=red_fill, font=red_font))
er = ilast + 3
ws.cell(row=er, column=2, value="Event-at-risk rule (composite)").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
ws.cell(row=er+1, column=2, value="atRisk = eventDate >= now AND ( hasOpenCriticalIssue OR (daysUntilEvent <= 30 AND readiness < 70 ) )").font = CODE_FONT
ws.merge_cells(start_row=er+1, start_column=2, end_row=er+1, end_column=7)
hr = er + 3
ehdrs = ["Event", "Event date", "Readiness", "Has open critical?", "Days until", "At risk?"]
for i, h in enumerate(ehdrs): ws.cell(row=hr, column=i+1, value=h)
style_header(ws, hr, 6)
events = [
    ("Karim & Nour — Garden Zaffa",       "=TODAY()+18", 62, "NO"),
    ("Omar & Salma — North Coast Reception","=TODAY()+40", 88, "NO"),
    ("Hany & Dahlia — Zamalek Wedding",   "=TODAY()-3",  100, "NO"),
    ("Reem & Tarek — Henna Night",        "=TODAY()+9",  71, "YES"),
]
r = hr + 1
for name, edate, ready, hascrit in events:
    ws.cell(row=r, column=2, value=name).font = B_FONT
    c = ws.cell(row=r, column=3, value=edate); c.number_format = DATEFMT; c.alignment = CTR
    ws.cell(row=r, column=4, value=ready).alignment = CTR
    c = ws.cell(row=r, column=5, value=hascrit); c.fill = INPUT_FILL; c.alignment = CTR; c.border = BORDER
    ws.cell(row=r, column=6, value=f'=IF($C{r}>=$C$4,CEILING($C{r}-$C$4,1),"")').alignment = CTR
    ws.cell(row=r, column=7, value=f'=IF(AND($C{r}>=$C$4,OR($E{r}="YES",AND($F{r}<>"",$F{r}<=$C$5,$D{r}<70))),"YES","NO")').alignment = CTR
    for col in range(2, 8):
        ws.cell(row=r, column=col).border = BORDER
        if col in (6,7): ws.cell(row=r, column=col).fill = CALC_FILL; ws.cell(row=r, column=col).font = B_FONT
    r += 1
elast = r - 1
ws.conditional_formatting.add(f"G{hr+1}:G{elast}", CellIsRule(operator="equal", formula=['"YES"'], fill=red_fill, font=red_font))
ws.conditional_formatting.add(f"E{hr+1}:E{elast}", CellIsRule(operator="equal", formula=['"YES"'], fill=red_fill, font=red_font))
cr2 = elast + 3
ws.cell(row=cr2, column=2, value="Exact source expressions").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
refs2 = [
    ("Open critical issue",          "i.priority === 'CRITICAL' && i.status !== 'RESOLVED'",                                          "api/dashboard/route.ts:132-134"),
    ("Open high/critical (readiness)","(i.priority === 'HIGH' || i.priority === 'CRITICAL') && i.status !== 'RESOLVED'",             "lib/readiness.ts:151-153"),
    ("Event has critical issue",     "issuesAll.some(i => i.eventId === e.id && i.priority === 'CRITICAL' && i.status !== 'RESOLVED')","api/dashboard/route.ts:105-107"),
    ("Event at risk",                "eventDate >= now && (hasCriticalIssue || (days <= 30 && readiness < 70))",                      "api/dashboard/route.ts:108"),
    ("Critical-issues criterion",    "score = openCritical.length === 0 ? 1 : 0  (weight 15 pts)",                                    "lib/readiness.ts:154,189"),
]
for i, (rule, expr, src) in enumerate(refs2):
    rr = cr2 + 1 + i
    ws.cell(row=rr, column=2, value=rule).font = B_FONT; ws.cell(row=rr, column=2).border = BORDER
    ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=6)
    c = ws.cell(row=rr, column=3, value=expr); c.font = CODE_FONT; c.alignment = WRAP; c.border = BORDER
    ws.merge_cells(start_row=rr, start_column=7, end_row=rr, end_column=8)
    ws.cell(row=rr, column=7, value=src).font = MUTED; ws.cell(row=rr, column=7).border = BORDER
ws.freeze_panes = "B9"

# ============ SHEET 4: Staff & Workload ============
ws = wb.create_sheet("Staff & Workload")
title_block(ws, "Staff & Workload",
            "Source: api/dashboard/route.ts:239 (my tasks), lib/readiness.ts:96-98 (run-sheet owners), api/dashboard/route.ts:248-253 (completion & avg readiness).", span=8)
widths(ws, {"A": 12, "B": 22, "C": 14, "D": 16, "E": 14, "F": 16, "G": 14, "H": 16})
ws["B4"] = "start of today"; ws["C4"] = "=TODAY()"; ws["C4"].number_format = DATEFMT; ws["C4"].fill = INPUT_FILL; ws["C4"].border = BORDER; ws["C4"].alignment = CTR
ws.cell(row=4, column=2).font = Font(name=FONT_NAME, size=10, bold=True); ws.cell(row=4, column=2).alignment = RIGHT
HR = 7
shdrs = ["Staff ID", "Name", "Role", "Total assigned", "Overdue", "Open critical", "Completed", "Completion %"]
for i, h in enumerate(shdrs): ws.cell(row=HR, column=i+1, value=h)
style_header(ws, HR, 8)
staff = [("u_amira","Amira Hassan","CEO"),("u_tarek","Tarek Mansour","FINANCE"),("u_omar","Omar Khaled","OPERATIONS"),("u_sara","Sara Adel","OPERATIONS")]
TASKS_START = 18
r = HR + 1
for sid, name, role in staff:
    ws.cell(row=r, column=1, value=sid).font = CODE_FONT
    ws.cell(row=r, column=2, value=name).font = B_FONT
    ws.cell(row=r, column=3, value=role).alignment = CTR
    rng = f"$B${TASKS_START}:$B${TASKS_START+200}"
    srng = f"$C${TASKS_START}:$C${TASKS_START+200}"  # priority
    sttng = f"$D${TASKS_START}:$D${TASKS_START+200}" # status
    dueng = f"$E${TASKS_START}:$E${TASKS_START+200}" # due
    ws.cell(row=r, column=4, value=f'=COUNTIFS({rng},$A{r})').alignment = CTR
    ws.cell(row=r, column=5, value=f'=COUNTIFS({rng},$A{r},{sttng},"<>COMPLETED",{dueng},"<"&$C$4)').alignment = CTR
    ws.cell(row=r, column=6, value=f'=COUNTIFS({rng},$A{r},{srng},"CRITICAL",{sttng},"<>COMPLETED")').alignment = CTR
    ws.cell(row=r, column=7, value=f'=COUNTIFS({rng},$A{r},{sttng},"COMPLETED")').alignment = CTR
    ws.cell(row=r, column=8, value=f'=IFERROR($G{r}/$D{r},0)').alignment = CTR
    ws.cell(row=r, column=8).number_format = PCTFMT
    for col in range(1, 9):
        ws.cell(row=r, column=col).border = BORDER
        if col >= 4: ws.cell(row=r, column=col).fill = CALC_FILL; ws.cell(row=r, column=col).font = B_FONT
    r += 1
slast = r - 1
ws.conditional_formatting.add(f"E{HR+1}:E{slast}", CellIsRule(operator="greaterThan", formula=["0"], fill=red_fill, font=red_font))
ws.conditional_formatting.add(f"F{HR+1}:F{slast}", CellIsRule(operator="greaterThan", formula=["0"], fill=red_fill, font=red_font))
tr = TASKS_START - 1
thdrs = ["Task ID", "Assignee", "Priority", "Status", "Due date", "Event"]
for i, h in enumerate(thdrs): ws.cell(row=tr, column=i+1, value=h)
style_header(ws, tr, 6)
tasks_data = [
    ("t01","u_omar","HIGH","TODO","=TODAY()-2","Karim & Nour"),
    ("t02","u_omar","CRITICAL","IN_PROGRESS","=TODAY()-1","Karim & Nour"),
    ("t03","u_omar","MEDIUM","COMPLETED","=TODAY()-5","Karim & Nour"),
    ("t04","u_omar","LOW","TODO","=TODAY()+3","Reem & Tarek"),
    ("t05","u_sara","HIGH","TODO","=TODAY()+1","Omar & Salma"),
    ("t06","u_sara","CRITICAL","TODO","=TODAY()-3","Omar & Salma"),
    ("t07","u_sara","MEDIUM","IN_PROGRESS","=TODAY()","Omar & Salma"),
    ("t08","u_sara","HIGH","COMPLETED","=TODAY()-4","Omar & Salma"),
    ("t09","u_tarek","MEDIUM","TODO","=TODAY()+5","Reem & Tarek"),
    ("t10","u_tarek","LOW","COMPLETED","=TODAY()-7","Hany & Dahlia"),
    ("t11","u_amira","HIGH","COMPLETED","=TODAY()-10","Hany & Dahlia"),
    ("t12","u_omar","MEDIUM","BLOCKED","=TODAY()-1","Karim & Nour"),
]
r = TASKS_START
for tid, assignee, pri, status, due, ev in tasks_data:
    ws.cell(row=r, column=1, value=tid).font = CODE_FONT
    ws.cell(row=r, column=2, value=assignee).font = CODE_FONT
    ws.cell(row=r, column=3, value=pri).alignment = CTR
    ws.cell(row=r, column=4, value=status).alignment = CTR
    c = ws.cell(row=r, column=5, value=due); c.number_format = DATEFMT; c.alignment = CTR
    ws.cell(row=r, column=6, value=ev).font = B_FONT
    for col in range(1, 7): ws.cell(row=r, column=col).border = BORDER
    r += 1
tlast = r - 1
rr = tlast + 3
ws.cell(row=rr, column=2, value="Run-sheet completion (per event)").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
ws.cell(row=rr+1, column=2, value="round(doneItems / totalItems * 100), or 0 if none").font = CODE_FONT
ws.merge_cells(start_row=rr+1, start_column=2, end_row=rr+1, end_column=6)
ws.cell(row=rr+3, column=2, value="Event").font = H_FONT; ws.cell(row=rr+3, column=2).fill = HEADER_FILL; ws.cell(row=rr+3, column=2).border = BORDER
ws.cell(row=rr+3, column=3, value="Run-sheet items").font = H_FONT; ws.cell(row=rr+3, column=3).fill = HEADER_FILL; ws.cell(row=rr+3, column=3).border = BORDER
ws.cell(row=rr+3, column=4, value="Done items").font = H_FONT; ws.cell(row=rr+3, column=4).fill = HEADER_FILL; ws.cell(row=rr+3, column=4).border = BORDER
ws.cell(row=rr+3, column=5, value="Completion %").font = H_FONT; ws.cell(row=rr+3, column=5).fill = HEADER_FILL; ws.cell(row=rr+3, column=5).border = BORDER
run_data = [("Karim & Nour", 24, 18), ("Omar & Salma", 12, 12), ("Reem & Tarek", 6, 2)]
for i, (ev, total, done) in enumerate(run_data):
    rrr = rr + 4 + i
    ws.cell(row=rrr, column=2, value=ev).font = B_FONT; ws.cell(row=rrr, column=2).border = BORDER
    c = ws.cell(row=rrr, column=3, value=total); c.fill = INPUT_FILL; c.alignment = CTR; c.border = BORDER
    c = ws.cell(row=rrr, column=4, value=done); c.fill = INPUT_FILL; c.alignment = CTR; c.border = BORDER
    c = ws.cell(row=rrr, column=5, value=f'=IFERROR(ROUND($D{rrr}/$C{rrr}*100,0),0)'); c.fill = CALC_FILL; c.alignment = CTR; c.border = BORDER; c.number_format = "0\"%\""
cr3 = rr + 4 + len(run_data) + 2
ws.cell(row=cr3, column=2, value="Exact source expressions").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
refs3 = [
    ("My tasks (Ops view)",        "tasksAll.filter(t => t.assignedToId === auth.id)", "api/dashboard/route.ts:239"),
    ("Run-sheet item has owner",   "r.responsibleId || r.responsibleName",            "lib/readiness.ts:96"),
    ("Owner share",                "items === 0 ? 0 : withOwner / items",             "lib/readiness.ts:98"),
    ("Run-sheet completion %",     "runItems.length === 0 ? 0 : Math.round(doneItems / runItems.length * 100)", "api/dashboard/route.ts:248-249"),
    ("Avg readiness",              "Math.round(sum(readiness) / events.length)",      "api/dashboard/route.ts:250-253"),
]
for i, (rule, expr, src) in enumerate(refs3):
    rrr = cr3 + 1 + i
    ws.cell(row=rrr, column=2, value=rule).font = B_FONT; ws.cell(row=rrr, column=2).border = BORDER
    ws.merge_cells(start_row=rrr, start_column=3, end_row=rrr, end_column=6)
    c = ws.cell(row=rrr, column=3, value=expr); c.font = CODE_FONT; c.alignment = WRAP; c.border = BORDER
    ws.merge_cells(start_row=rrr, start_column=7, end_row=rrr, end_column=8)
    ws.cell(row=rrr, column=7, value=src).font = MUTED; ws.cell(row=rrr, column=7).border = BORDER
ws.freeze_panes = "B8"

# ============ SHEET 5: Readiness Score ============
ws = wb.create_sheet("Readiness Score")
title_block(ws, "Event Readiness Score — Live Calculator",
            "Source: lib/readiness.ts (full model) + lib/constants.ts:183-190 (weights). Change the amber inputs and watch the 100-point score update.", span=7)
widths(ws, {"A": 4, "B": 36, "C": 18, "D": 10, "E": 14, "F": 12, "G": 44})
ws.cell(row=4, column=2, value="INPUTS (edit amber cells)").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
inputs = [
    ("Vendors engaged (total)",         8),
    ("Vendors confirmed or beyond",     5),
    ("Critical tasks (HIGH/CRITICAL) total", 6),
    ("Critical tasks completed",        4),
    ("Run-sheet items (total)",         10),
    ("Run-sheet items with an owner",   8),
    ("Documents uploaded",              2),
    ("Open high/critical issues",       1),
    ("Event status",                    "CONFIRMED"),
    ("Has venue recorded? (YES/NO)",    "YES"),
    ("Venue vendor confirmed? (YES/NO)","YES"),
]
r = 5
for lbl, val in inputs:
    ws.cell(row=r, column=2, value=lbl).font = B_FONT; ws.cell(row=r, column=2).alignment = RIGHT
    c = ws.cell(row=r, column=3, value=val); c.fill = INPUT_FILL; c.border = BORDER; c.alignment = CTR; c.font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
    r += 1
dv = DataValidation(type="list", formula1='"LEAD,PROPOSAL_SENT,CONFIRMED,IN_PREPARATION,READY,COMPLETED,ARCHIVED"', allow_blank=False)
ws.add_data_validation(dv); dv.add("C13")
dv2 = DataValidation(type="list", formula1='"YES,NO"', allow_blank=False)
ws.add_data_validation(dv2); dv2.add("C14"); dv2.add("C15")
cr = 17
ws.cell(row=cr, column=2, value="CRITERIA (live)").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
hdrs = ["#", "Criterion", "Weight", "Score (0-1)", "Points", "Met?", "Detail"]
for i, h in enumerate(hdrs): ws.cell(row=cr+1, column=i+1, value=h)
style_header(ws, cr+1, 7)
crit = [
    ("1", "Required vendors confirmed", 25, '=IFERROR(IF($C$5=0,0,$C$6/$C$5),0)',
     '=IF($C$5>0,"Confirmed "&$C$6&" of "&$C$5,"No vendors engaged yet")'),
    ("2", "Critical tasks completed", 20, '=IFERROR(IF($C$7=0,1,$C$8/$C$7),1)',
     '=IF($C$7=0,"No critical-priority tasks open",$C$8&" of "&$C$7&" high/critical done")'),
    ("3", "Run sheet prepared", 15, '=ROUND(MIN($C$9/6,1)*(0.6+0.4*IFERROR($C$10/$C$9,0)),2)',
     '=IF($C$9=0,"Run sheet is empty",$C$9&" items, "&($C$9-$C$10)&" missing an owner")'),
    ("4", "Venue confirmed", 15,
     '=IF(OR($C$15="YES",AND(OR($C$13="CONFIRMED",$C$13="IN_PREPARATION",$C$13="READY",$C$13="COMPLETED"),$C$14="YES")),1,0)',
     '=IF($C$15="YES","Venue vendor confirmed",IF(AND(OR($C$13="CONFIRMED",$C$13="IN_PREPARATION",$C$13="READY",$C$13="COMPLETED"),$C$14="YES"),"Event confirmed with venue recorded","Venue not confirmed yet"))'),
    ("5", "Required documents available", 10, '=MIN($C$11/2,1)',
     '=$C$11&" document"&IF($C$11=1,"","s")&" uploaded"'),
    ("6", "Critical issues resolved", 15, '=IF($C$12=0,1,0)',
     '=IF($C$12=0,"No open high/critical issues",$C$12&" open high/critical issue"&IF($C$12=1,"","s"))'),
]
r = cr + 2
for num, label, weight, score_fml, detail_fml in crit:
    ws.cell(row=r, column=1, value=num).alignment = CTR
    ws.cell(row=r, column=2, value=label).font = B_FONT
    ws.cell(row=r, column=3, value=weight).alignment = CTR
    sc = ws.cell(row=r, column=4, value=score_fml); sc.fill = CALC_FILL; sc.alignment = CTR; sc.number_format = "0.00"
    ws.cell(row=r, column=5, value=f'=$C{r}*$D{r}').fill = CALC_FILL; ws.cell(row=r, column=5).alignment = CTR; ws.cell(row=r, column=5).number_format = "0.0"
    ws.cell(row=r, column=6, value=f'=IF($D{r}>=1,"YES","NO")').alignment = CTR
    ws.cell(row=r, column=7, value=detail_fml).font = MUTED; ws.cell(row=r, column=7).alignment = WRAP
    for col in range(1, 8): ws.cell(row=r, column=col).border = BORDER
    r += 1
cfirst = cr + 2; clast = r - 1
tr_ = clast + 2
ws.cell(row=tr_, column=2, value="TOTAL READINESS SCORE").font = Font(name=FONT_NAME, size=12, bold=True, color=NEUTRAL_900); ws.cell(row=tr_, column=2).alignment = RIGHT
c = ws.cell(row=tr_, column=3, value=f'=ROUND(SUMPRODUCT($C${cfirst}:$C${clast},$D${cfirst}:$D${clast}),0)')
c.font = Font(name=FONT_NAME, size=14, bold=True, color=NEUTRAL_900); c.alignment = CTR; c.fill = PatternFill("solid", fgColor="FFF3B0"); c.border = BORDER
ws.cell(row=tr_, column=4, value="/ 100").font = MUTED; ws.cell(row=tr_, column=4).alignment = LEFT
ws.cell(row=tr_+1, column=2, value="Band").font = Font(name=FONT_NAME, size=11, bold=True); ws.cell(row=tr_+1, column=2).alignment = RIGHT
c = ws.cell(row=tr_+1, column=3, value=f'=IF($C${tr_}>=85,"On track",IF($C${tr_}>=70,"Attention","At risk"))')
c.alignment = CTR; c.font = Font(name=FONT_NAME, size=12, bold=True); c.fill = CALC_FILL; c.border = BORDER
ws.conditional_formatting.add(f"C{tr_+1}", CellIsRule(operator="equal", formula=['"At risk"'], fill=red_fill, font=Font(name=FONT_NAME, size=12, bold=True, color=ACCENT_NEGATIVE)))
ws.conditional_formatting.add(f"C{tr_+1}", CellIsRule(operator="equal", formula=['"Attention"'], fill=amber_fill, font=Font(name=FONT_NAME, size=12, bold=True, color=ACCENT_WARNING)))
ws.conditional_formatting.add(f"C{tr_+1}", CellIsRule(operator="equal", formula=['"On track"'], fill=green_fill, font=Font(name=FONT_NAME, size=12, bold=True, color=ACCENT_POSITIVE)))
wr = tr_ + 4
ws.cell(row=wr, column=2, value="Weights (from constants.ts)").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
wref = [("vendorsConfirmed",25),("criticalTasks",20),("runSheet",15),("venueConfirmed",15),("documents",10),("criticalIssues",15)]
for i,(k,v) in enumerate(wref):
    rrr = wr+1+i
    ws.cell(row=rrr, column=2, value=k).font = CODE_FONT
    ws.cell(row=rrr, column=3, value=v).alignment = CTR
ws.cell(row=wr+8, column=2, value="Band thresholds: >=85 On track | >=70 Attention | <70 At risk").font = MUTED
ws.merge_cells(start_row=wr+8, start_column=2, end_row=wr+8, end_column=7)
ws.cell(row=wr+9, column=2, value="Source: src/lib/readiness.ts (computeReadiness) | src/lib/constants.ts:183-190 (READINESS_WEIGHTS)").font = MUTED
ws.merge_cells(start_row=wr+9, start_column=2, end_row=wr+9, end_column=7)
ws.freeze_panes = "B5"

# ============ SHEET 6: Finance Formulas ============
ws = wb.create_sheet("Finance Formulas")
title_block(ws, "Event Financial Formulas",
            "Source: lib/finance.ts (eventFinancials). Ledger is immutable — reversals are negative-amount rows that net out automatically.", span=7)
widths(ws, {"A": 4, "B": 32, "C": 16, "D": 16, "E": 16, "F": 14, "G": 28})
ws.cell(row=4, column=2, value="EVENT INPUTS").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
ws.cell(row=5, column=2, value="Contract value").font = B_FONT; ws.cell(row=5, column=2).alignment = RIGHT
c = ws.cell(row=5, column=3, value=150000); c.fill = INPUT_FILL; c.border = BORDER; c.number_format = NUMFMT; c.alignment = CTR
ws.cell(row=6, column=2, value="Additional revenue").font = B_FONT; ws.cell(row=6, column=2).alignment = RIGHT
c = ws.cell(row=6, column=3, value=20000); c.fill = INPUT_FILL; c.border = BORDER; c.number_format = NUMFMT; c.alignment = CTR
ws.cell(row=7, column=2, value="Currency").font = B_FONT; ws.cell(row=7, column=2).alignment = RIGHT
c = ws.cell(row=7, column=3, value="EGP"); c.fill = INPUT_FILL; c.border = BORDER; c.alignment = CTR
# Vendor contracts (planned-expenses source — matches app's eventVendor.contractAmount, lib/finance.ts:48)
ws.cell(row=4, column=9, value="VENDOR CONTRACTS").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
ws.cell(row=5, column=9, value="Vendor").font = H_FONT; ws.cell(row=5, column=9).fill = HEADER_FILL; ws.cell(row=5, column=9).border = BORDER
ws.cell(row=5, column=10, value="Contract amount").font = H_FONT; ws.cell(row=5, column=10).fill = HEADER_FILL; ws.cell(row=5, column=10).border = BORDER
vc = [("Venue",35000),("Photographer",18000),("Caterer",12000),("Florist",8000),("Transport",5000)]
for i,(n,a) in enumerate(vc):
    rr = 6+i
    ws.cell(row=rr, column=9, value=n).font = B_FONT; ws.cell(row=rr, column=9).border = BORDER
    c = ws.cell(row=rr, column=10, value=a); c.fill = INPUT_FILL; c.number_format = NUMFMT; c.alignment = CTR; c.border = BORDER
ws.column_dimensions["I"].width = 18; ws.column_dimensions["J"].width = 18
HR = 10
phdrs = ["#", "Direction", "Kind", "Status", "Amount", "Due date", "Note"]
for i, h in enumerate(phdrs): ws.cell(row=HR, column=i+1, value=h)
style_header(ws, HR, 7)
pays = [
    ("INFLOW","CLIENT_PAYMENT","COMPLETED", 60000, "=TODAY()-30","Deposit"),
    ("INFLOW","CLIENT_PAYMENT","COMPLETED", 60000, "=TODAY()-10","2nd installment"),
    ("INFLOW","CLIENT_PAYMENT","PENDING",   50000, "=TODAY()+5", "Final balance"),
    ("OUTFLOW","VENDOR_PAYMENT","COMPLETED",35000, "=TODAY()-20","Venue"),
    ("OUTFLOW","VENDOR_PAYMENT","COMPLETED",18000, "=TODAY()-15","Photographer"),
    ("OUTFLOW","VENDOR_PAYMENT","PENDING",  12000, "=TODAY()-2", "Caterer deposit (overdue)"),
    ("OUTFLOW","EXPENSE","COMPLETED",       5000,  "=TODAY()-5", "Transport"),
    ("OUTFLOW","VENDOR_PAYMENT","COMPLETED",-35000, "=TODAY()-3", "Reversal: venue refund"),
]
r = HR + 1
for i,(d,k,s,amt,due,note) in enumerate(pays, 1):
    ws.cell(row=r, column=1, value=i).alignment = CTR
    ws.cell(row=r, column=2, value=d).font = CODE_FONT
    ws.cell(row=r, column=3, value=k).font = CODE_FONT
    ws.cell(row=r, column=4, value=s).alignment = CTR
    c = ws.cell(row=r, column=5, value=amt); c.number_format = NUMFMT; c.alignment = RIGHT
    c2 = ws.cell(row=r, column=6, value=due); c2.number_format = DATEFMT; c2.alignment = CTR
    ws.cell(row=r, column=7, value=note).font = B_FONT
    for col in range(1,8): ws.cell(row=r, column=col).border = BORDER
    r += 1
plast = r - 1
OR_ = plast + 3
ws.cell(row=OR_, column=2, value="COMPUTED FINANCIALS (live)").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
DR = f"$E${HR+1}:$E${plast}"; DIRR = f"$B${HR+1}:$B${plast}"; KR = f"$C${HR+1}:$C${plast}"; SR = f"$D${HR+1}:$D${plast}"
outs = [
    ("Total revenue",                f"=$C$5+$C$6", NUMFMT),
    ("Collected from client",        f'=SUMIFS({DR},{DIRR},"INFLOW",{KR},"CLIENT_PAYMENT",{SR},"COMPLETED")', NUMFMT),
    ("Outstanding client balance",   f"=MAX(0,$C${OR_+1}-$C${OR_+2})", NUMFMT),
    ("Planned expenses (vendor contracts)", f"=SUM($J$6:$J$10)", NUMFMT),
    ("Actual expenses (completed outflows)", f'=SUMIFS({DR},{DIRR},"OUTFLOW",{SR},"COMPLETED")', NUMFMT),
    ("Remaining budget",             f"=$C${OR_+4}-$C${OR_+5}", NUMFMT),
    ("Gross estimated profit",       f"=$C${OR_+1}-$C${OR_+4}", NUMFMT),
    ("Actual profit",                f"=$C${OR_+2}-$C${OR_+5}", NUMFMT),
    ("Profit margin %",              f"=IFERROR($C${OR_+7}/$C${OR_+1}*100,0)", "0.0\"%\""),
]
for i,(lbl,fml,fmt) in enumerate(outs):
    rr = OR_ + 1 + i
    ws.cell(row=rr, column=2, value=lbl).font = B_FONT; ws.cell(row=rr, column=2).alignment = RIGHT
    c = ws.cell(row=rr, column=3, value=fml); c.fill = CALC_FILL; c.border = BORDER; c.font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900); c.alignment = RIGHT; c.number_format = fmt
ovr = OR_ + len(outs) + 3
ws.cell(row=ovr, column=2, value="Overdue payment rule").font = Font(name=FONT_NAME, size=11, bold=True, color=NEUTRAL_900)
ws.cell(row=ovr+1, column=2, value="p.status === 'PENDING' && p.dueDate !== null && p.dueDate < now").font = CODE_FONT
ws.merge_cells(start_row=ovr+1, start_column=2, end_row=ovr+1, end_column=7)
ws.cell(row=ovr+2, column=2, value="Overdue payments in the table above:").font = B_FONT
FR = f"$F${HR+1}:$F${plast}"
c = ws.cell(row=ovr+2, column=3, value=f'=SUMPRODUCT(({SR}="PENDING")*({FR}<>"")*({FR}<TODAY()))')
c.alignment = CTR; c.fill = CALC_FILL; c.border = BORDER; c.font = Font(name=FONT_NAME, size=11, bold=True, color=ACCENT_NEGATIVE)
fr_ = ovr + 5
ws.cell(row=fr_, column=2, value="Source: src/lib/finance.ts — eventFinancials()").font = MUTED
ws.merge_cells(start_row=fr_, start_column=2, end_row=fr_, end_column=7)
ws.freeze_panes = "B11"

# ============ SHEET 7: Formula Reference ============
ws = wb.create_sheet("Formula Reference")
title_block(ws, "Formula Reference — Master",
            "Every business formula in Potcha Ops: the plain-English rule, the exact code expression, and where it lives in source.", span=5)
widths(ws, {"A": 4, "B": 18, "C": 30, "D": 62, "E": 30})
hdrs = ["#", "Domain", "Formula", "Code expression", "Source"]
for i, h in enumerate(hdrs): ws.cell(row=4, column=i+1, value=h)
style_header(ws, 4, 5)
all_refs = [
    ("Task delay", "Task isOverdue (display flag)", "t.dueDate !== null && t.dueDate < now && t.status !== 'COMPLETED'", "api/tasks/route.ts:60"),
    ("Task delay", "Overdue task (dashboard)", "t.status !== 'COMPLETED' && t.dueDate && t.dueDate < startOfToday", "api/dashboard/route.ts:126-128"),
    ("Task delay", "Task due today", "t.status !== 'COMPLETED' && t.dueDate >= startOfToday && t.dueDate <= endOfToday", "api/dashboard/route.ts:129-131"),
    ("Task delay", "Overdue query filter", "where.dueDate = { lt: startOfToday }", "api/tasks/route.ts:40"),
    ("Task delay", "Days until event", "Math.ceil((eventDate - now) / 86400000)", "api/dashboard/route.ts:42"),
    ("Task delay", "Payment overdue", "p.status === 'PENDING' && p.dueDate !== null && p.dueDate < now", "lib/finance.ts:30"),
    ("Task delay", "Payment due <= 14 days", "p.dueDate >= startOfToday && p.dueDate <= now + 14*86400000", "api/dashboard/route.ts:121-122"),
    ("Critical issues", "Open critical issue", "i.priority === 'CRITICAL' && i.status !== 'RESOLVED'", "api/dashboard/route.ts:132-134"),
    ("Critical issues", "Open high/critical", "(i.priority === 'HIGH' || i.priority === 'CRITICAL') && i.status !== 'RESOLVED'", "lib/readiness.ts:151-153"),
    ("Critical issues", "Event has critical issue", "issuesAll.some(i => i.eventId === e.id && i.priority === 'CRITICAL' && i.status !== 'RESOLVED')", "api/dashboard/route.ts:105-107"),
    ("Critical issues", "Event at risk", "eventDate >= now && (hasCriticalIssue || (daysUntil <= 30 && readiness < 70))", "api/dashboard/route.ts:108"),
    ("Critical issues", "Critical-issues criterion score", "score = openCritical.length === 0 ? 1 : 0  (weight 15)", "lib/readiness.ts:154"),
    ("Staff", "My tasks (Ops view)", "tasksAll.filter(t => t.assignedToId === auth.id)", "api/dashboard/route.ts:239"),
    ("Staff", "Run-sheet item has owner", "r.responsibleId || r.responsibleName", "lib/readiness.ts:96"),
    ("Staff", "Owner share", "items === 0 ? 0 : withOwner / items", "lib/readiness.ts:98"),
    ("Staff", "Run-sheet completion %", "runItems.length === 0 ? 0 : Math.round(doneItems / runItems.length * 100)", "api/dashboard/route.ts:248-249"),
    ("Staff", "Avg readiness", "events.length === 0 ? 0 : Math.round(sum(readiness) / events.length)", "api/dashboard/route.ts:250-253"),
    ("Staff", "Vendor confirmations outstanding", "count where status in [NOT_CONTACTED, CONTACTED, NEGOTIATING, DEPOSIT_REQUIRED]", "api/dashboard/route.ts:135-137"),
    ("Readiness", "Vendors confirmed score", "total === 0 ? 0 : confirmed / total  (weight 25)", "lib/readiness.ts:57-72"),
    ("Readiness", "Critical tasks score", "critical.length === 0 ? 1 : done / critical.length  (weight 20)", "lib/readiness.ts:77-90"),
    ("Readiness", "Run-sheet score", "round(min(items/6,1) * (0.6 + 0.4*ownerShare))  (weight 15)", "lib/readiness.ts:94-110"),
    ("Readiness", "Venue score", "venueVendorConfirmed || (statusConfirmed && hasVenue) ? 1 : 0  (weight 15)", "lib/readiness.ts:114-133"),
    ("Readiness", "Documents score", "min(documents / 2, 1)  (weight 10)", "lib/readiness.ts:136-146"),
    ("Readiness", "Total score", "Math.round(sum(weight * score))", "lib/readiness.ts:168-169"),
    ("Readiness", "Band", "score >= 85 ? on_track : score >= 70 ? attention : at_risk", "lib/readiness.ts:193"),
    ("Finance", "Total revenue", "contractValue + additionalRevenue", "lib/finance.ts:39"),
    ("Finance", "Collected from client", "sum(COMPLETED INFLOW CLIENT_PAYMENT) — reversals net out", "lib/finance.ts:41-45"),
    ("Finance", "Outstanding balance", "Math.max(0, totalRevenue - collectedFromClient)", "lib/finance.ts:57"),
    ("Finance", "Planned expenses", "sum(eventVendor.contractAmount)", "lib/finance.ts:48"),
    ("Finance", "Actual expenses", "sum(COMPLETED OUTFLOW)", "lib/finance.ts:50-54"),
    ("Finance", "Remaining budget", "plannedExpenses - actualExpenses", "lib/finance.ts:58"),
    ("Finance", "Gross estimated profit", "totalRevenue - plannedExpenses", "lib/finance.ts:59"),
    ("Finance", "Actual profit", "collectedFromClient - actualExpenses", "lib/finance.ts:60"),
    ("Finance", "Profit margin %", "totalRevenue > 0 ? (grossEstimatedProfit / totalRevenue) * 100 : 0", "lib/finance.ts:61"),
    ("Finance", "round2", "Math.round((n + Number.EPSILON) * 100) / 100", "lib/finance.ts:22-24"),
]
r = 5
for i, (domain, name, expr, src) in enumerate(all_refs, 1):
    ws.cell(row=r, column=1, value=i).alignment = CTR
    ws.cell(row=r, column=2, value=domain).font = Font(name=FONT_NAME, size=10, bold=True, color=NEUTRAL_900)
    ws.cell(row=r, column=3, value=name).font = B_FONT; ws.cell(row=r, column=3).alignment = WRAP
    c = ws.cell(row=r, column=4, value=expr); c.font = CODE_FONT; c.alignment = WRAP
    ws.cell(row=r, column=5, value=src).font = MUTED
    for col in range(1, 6): ws.cell(row=r, column=col).border = BORDER
    r += 1
ws.auto_filter.ref = f"A4:E{r-1}"
ws.freeze_panes = "A5"

out = "/home/z/my-project/Potcha-Ops-Business-Formulas.xlsx"
wb.save(out)
print(f"SAVED: {out}")
print(f"Sheets: {wb.sheetnames}")
