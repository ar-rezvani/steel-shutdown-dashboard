"""
Generate a fake (but realistic) steel-plant shutdown repair dataset for the Power BI dashboard.

Output: data/ShutdownRepair_SampleData.xlsx   (+ docs/04-validation.md with expected KPI values)

Usage:  python scripts/generate_sample_data.py
Deterministic: same seed -> same data, so expected KPIs in docs/04-validation.md stay valid.
"""
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SEED = 42
N_ACTIVITIES = 120
SD_START = date(2026, 10, 17)   # shutdown start (fake)
SD_DAYS = 14                    # shutdown length in days
HOURS_PER_PERSON = 10           # working hours per person per day

ROOT = Path(__file__).resolve().parent.parent
OUT_XLSX = ROOT / "data" / "ShutdownRepair_SampleData.xlsx"
OUT_VALIDATION = ROOT / "docs" / "04-validation.md"

random.seed(SEED)

# --------------------------------------------------------------------------- dimensions
AREAS = [
    ("A01", "کوره قوس الکتریکی", "EAF"),
    ("A02", "کوره پاتیلی", "Ladle Furnace"),
    ("A03", "ریخته‌گری مداوم", "CCM"),
    ("A04", "واحد تصفیه گاز", "Gas Cleaning"),
    ("A05", "سیستم آب خنک‌کاری", "Cooling Water"),
    ("A06", "جرثقیل‌های سقفی", "Overhead Cranes"),
    ("A07", "پست برق و MCC", "Power & MCC"),
]
DISCIPLINES = [
    ("D01", "مکانیک", "Mechanical"),
    ("D02", "برق", "Electrical"),
    ("D03", "ابزار دقیق", "Instrumentation"),
    ("D04", "جوشکاری", "Welding"),
    ("D05", "نسوزکاری", "Refractory"),
    ("D06", "داربست", "Scaffolding"),
    ("D07", "عمران", "Civil"),
]
CONTRACTORS = [
    ("C01", "تیم داخلی نگهداری و تعمیرات", "Internal"),
    ("C02", "پیمانکار الف", "External"),
    ("C03", "پیمانکار ب", "External"),
    ("C04", "پیمانکار ج", "External"),
    ("C05", "پیمانکار د", "External"),
]
REASONS = [  # id, name FA, category, weight
    ("R01", "کمبود قطعه یدکی / مواد", "Material", 24),
    ("R02", "کمبود نیروی انسانی", "Manpower", 14),
    ("R03", "تأخیر صدور مجوز کار", "Permit", 12),
    ("R04", "تأخیر پیمانکار", "Contractor", 14),
    ("R05", "کشف کار اضافه (Extra Scope)", "Scope", 12),
    ("R06", "مشکل دسترسی / داربست", "Access", 8),
    ("R07", "خرابی تجهیزات کمکی (جرثقیل/ابزار)", "Equipment", 6),
    ("R08", "محدودیت زمان توقف", "Time", 10),
]

TASKS = {
    "D01": ["تعویض غلتک‌ها", "بازرسی و تعویض یاتاقان", "آلاینمنت گیربکس", "تعویض شیلنگ‌های هیدرولیک",
            "اورهال پمپ", "تعویض سیل و کوپلینگ", "سرویس سیلندر هیدرولیک", "تعویض فن و دمپر"],
    "D02": ["تعویض کابل قدرت", "تست رله‌های حفاظتی", "سرویس تابلو MCC", "تعویض الکتروموتور",
            "تست عایقی ترانسفورماتور", "بازبینی سیستم زمین", "تعویض کنتاکتور و بریکر"],
    "D03": ["کالیبراسیون ترانسمیتر فشار", "تعویض ترموکوپل", "تست حلقه‌های کنترلی", "سرویس سنسور سطح",
            "کالیبراسیون فلومتر", "تعویض شیر کنترل"],
    "D04": ["جوشکاری و ترمیم سازه", "تعویض پنل آب‌خنک", "ترمیم لوله‌های فرآیندی", "جوشکاری ساپورت‌ها",
            "تعویض لوله‌های خنک‌کاری"],
    "D05": ["نسوزکاری دیواره", "ترمیم آستر پاتیل", "تعویض آجر کف", "گانیتینگ و ترمیم نسوز", "خشک‌کردن و پخت نسوز"],
    "D06": ["نصب داربست", "جابه‌جایی داربست", "برچیدن داربست", "داربست آویز"],
    "D07": ["ترمیم فونداسیون", "تعمیر کف‌سازی", "ترمیم کانال و فاضلاب", "تقویت سازه بتنی"],
}
MH_RANGE = {"D01": (40, 400), "D02": (30, 250), "D03": (20, 150), "D04": (40, 300),
            "D05": (100, 600), "D06": (30, 200), "D07": (40, 250)}
CONTRACTOR_BY_DISC = {"D01": ["C01", "C02"], "D02": ["C01", "C03"], "D03": ["C01", "C03"],
                      "D04": ["C02", "C04"], "D05": ["C05"], "D06": ["C04"], "D07": ["C02", "C04"]}
PRIORITIES = ["Critical", "High", "Normal"]


# --------------------------------------------------------------------------- helpers
def gregorian_to_jalali(gy, gm, gd):
    """Standard Gregorian -> Jalali conversion (no external dependency)."""
    g_d_m = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
    gy2 = gy + 1 if gm > 2 else gy
    days = (355666 + 365 * gy + (gy2 + 3) // 4 - (gy2 + 99) // 100 + (gy2 + 399) // 400
            + gd + g_d_m[gm - 1])
    jy = -1595 + 33 * (days // 12053)
    days %= 12053
    jy += 4 * (days // 1461)
    days %= 1461
    if days > 365:
        jy += (days - 1) // 365
        days = (days - 1) % 365
    if days < 186:
        jm, jd = 1 + days // 31, 1 + days % 31
    else:
        jm, jd = 7 + (days - 186) // 30, 1 + (days - 186) % 30
    return jy, jm, jd


def distribute(total, weights):
    """Split `total` by weights into 2-decimals values that sum EXACTLY to total."""
    s = sum(weights)
    if s == 0 or total == 0:
        return [0.0] * len(weights)
    parts = [round(total * w / s, 2) for w in weights]
    last = max(i for i, w in enumerate(weights) if w > 0)
    parts[last] = round(parts[last] + (round(total, 2) - round(sum(parts), 2)), 2)
    return parts


# --------------------------------------------------------------------------- activities
activities = []
for i in range(1, N_ACTIVITIES + 1):
    area = random.choice(AREAS)
    disc = random.choices(DISCIPLINES, weights=[28, 20, 14, 14, 9, 9, 6])[0]
    d_id = disc[0]
    task = random.choice(TASKS[d_id])
    lo, hi = MH_RANGE[d_id]
    planned = round(random.randint(lo, hi) / 5) * 5
    dur = random.randint(1, 6)
    start_off = random.randint(0, SD_DAYS - dur)
    p_start = SD_START + timedelta(days=start_off)
    p_end = p_start + timedelta(days=dur - 1)

    r = random.random()
    if r < 0.58:
        progress = 1.0
    elif r < 0.78:
        progress = round(random.uniform(0.15, 0.90), 2)
    else:
        progress = 0.0

    earned = round(planned * progress, 2)
    actual = round(earned * random.uniform(0.85, 1.35), 2) if progress > 0 else 0.0
    reason = ""
    if progress < 1.0:
        reason = random.choices([x[0] for x in REASONS], weights=[x[3] for x in REASONS])[0]

    activities.append(dict(
        id=f"SD-{i:03d}", desc=f"{task} - {area[1]}", area=area[0], disc=d_id,
        contractor=random.choice(CONTRACTOR_BY_DISC[d_id]),
        priority=random.choices(PRIORITIES, weights=[25, 35, 40])[0],
        start=p_start, end=p_end, planned=planned, actual=actual, progress=progress,
        earned=earned, reason=reason, dur=dur,
    ))

# --------------------------------------------------------------------------- daily fact
daily = []
for a in activities:
    days = [a["start"] + timedelta(days=k) for k in range(a["dur"])]
    planned_parts = distribute(a["planned"], [1] * a["dur"])
    k_work = 0 if a["progress"] == 0 else max(1, round(a["dur"] * a["progress"]))
    w = [random.uniform(0.6, 1.4) if idx < k_work else 0 for idx in range(a["dur"])]
    actual_parts = distribute(a["actual"], w)
    earned_parts = distribute(a["earned"], w)
    for d, p, ac, ea in zip(days, planned_parts, actual_parts, earned_parts):
        daily.append((d, a["id"], p, ac, ea))

# --------------------------------------------------------------------------- resources
planned_by_disc_day = {}
for d, aid, p, _, _ in daily:
    disc = next(x["disc"] for x in activities if x["id"] == aid)
    planned_by_disc_day[(disc, d)] = planned_by_disc_day.get((disc, d), 0) + p
avg_by_disc = {}
for disc, *_ in DISCIPLINES:
    vals = [v for (dd, _), v in planned_by_disc_day.items() if dd == disc]
    avg_by_disc[disc] = sum(vals) / SD_DAYS

resources = []
for k in range(SD_DAYS):
    d = SD_START + timedelta(days=k)
    for disc, *_ in DISCIPLINES:
        need = avg_by_disc[disc] * random.uniform(1.0, 1.35)
        headcount = max(2, round(need / HOURS_PER_PERSON))
        resources.append((d, disc, headcount, HOURS_PER_PERSON))

# --------------------------------------------------------------------------- date dimension (Jalali)
MONTHS_FA = ["فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
             "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"]
WEEKDAYS_FA = ["شنبه", "یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه", "پنج‌شنبه", "جمعه"]
assert gregorian_to_jalali(2026, 3, 21) == (1405, 1, 1), "Jalali conversion sanity check failed"

date_rows = []
d = date(2026, 3, 21)               # 1 Farvardin 1405
while d <= date(2027, 3, 20):       # 29 Esfand 1405
    jy, jm, jd = gregorian_to_jalali(d.year, d.month, d.day)
    wd = (d.weekday() + 2) % 7      # Saturday = 0
    date_rows.append((d, d.year, d.month, d.strftime("%B"), d.day,
                      f"{jy}/{jm:02d}/{jd:02d}", jy, jm, MONTHS_FA[jm - 1], jd,
                      f"{jy}-{jm:02d}", wd, WEEKDAYS_FA[wd], 1 if wd == 6 else 0))
    d += timedelta(days=1)

# --------------------------------------------------------------------------- write workbook
FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
BODY = Font(name=FONT, size=10)
INPUT = Font(name=FONT, size=10, color="0000FF")      # blue = user input
CALC = Font(name=FONT, size=10, color="000000")       # black = formula


def write_sheet(wb, title, headers, rows, widths, fmts=None, fonts=None, rtl=True):
    ws = wb.create_sheet(title)
    ws.sheet_view.rightToLeft = rtl
    ws.append(headers)
    for c in ws[1]:
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in rows:
        ws.append(list(r))
    fmts, fonts = fmts or {}, fonts or {}
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = fonts.get(c.column_letter, BODY)
            if c.column_letter in fmts:
                c.number_format = fmts[c.column_letter]
    for i, wdt in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = wdt
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return ws


wb = Workbook()
wb.remove(wb.active)

# README sheet -------------------------------------------------------------
rd = wb.create_sheet("README")
rd.sheet_view.rightToLeft = True
rd.column_dimensions["A"].width = 26
rd.column_dimensions["B"].width = 95
lines = [
    ("فایل ورودی داشبورد گزارش تعمیرات توقف (Shutdown) - واحد فولادسازی", None),
    ("هشدار", "تمام داده‌های این فایل فرضی (Fake) است و فقط برای توسعه و تست داشبورد ساخته شده."),
    ("", ""),
    ("راهنمای رنگ سلول‌ها", None),
    ("متن آبی", "ورودی کاربر؛ این ستون‌ها را به‌روز کنید (ActualMH، Progress، ReasonID و ...)."),
    ("متن مشکی", "فرمول یا جدول مرجع؛ دستی تغییر ندهید (EarnedMH، Status، StatusFA، AvailableMH)."),
    ("", ""),
    ("شیت‌ها", None),
    ("Activities", "فهرست فعالیت‌ها: نفر‑ساعت برنامه‌ای، نفر‑ساعت واقعی، درصد پیشرفت، علت عدم انجام. (نمونه: ردیف اول)"),
    ("Daily_Progress", "ریز روزانه هر فعالیت: PlannedMH_Day / ActualMH_Day / EarnedMH_Day (برای منحنی S)."),
    ("Resources", "نفر‑ساعت در اختیار هر تخصص در هر روز (Headcount × Hours)."),
    ("Dim_*", "جداول مرجع: Area, Discipline, Contractor, Reason, Date (با تاریخ شمسی)."),
    ("", ""),
    ("تعاریف", None),
    ("PlannedMH", "نفر‑ساعت برنامه‌ریزی‌شده فعالیت."),
    ("ActualMH", "نفر‑ساعت واقعاً مصرف‌شده."),
    ("EarnedMH", "نفر‑ساعت کسب‌شده = PlannedMH × Progress (فرمول)."),
    ("Status", "Done اگر Progress=100٪ ؛ In Progress اگر 0<Progress<100٪ ؛ Not Done اگر Progress=0 (فرمول)."),
    ("", ""),
    ("روش به‌روزرسانی", None),
    ("1", "مقادیر ستون‌های آبی را در Activities / Daily_Progress / Resources جایگزین کنید."),
    ("2", "فایل را ذخیره و ببندید؛ در Power BI روی Refresh بزنید."),
    ("3", "نام شیت‌ها و عنوان ستون‌ها را تغییر ندهید."),
]
for a, b in lines:
    rd.append([a, b] if b is not None else [a])
    row = rd[rd.max_row]
    if b is None:
        row[0].font = Font(name=FONT, bold=True, size=11, color="1F4E79")
    else:
        row[0].font = Font(name=FONT, bold=True, size=10)
        row[1].font = BODY
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
rd["A5"].font = Font(name=FONT, bold=True, size=10, color="0000FF")

# Activities ---------------------------------------------------------------
act_headers = ["ActivityID", "Description", "AreaID", "DisciplineID", "ContractorID", "Priority",
               "PlannedStart", "PlannedEnd", "PlannedMH", "ActualMH", "Progress",
               "EarnedMH", "Status", "StatusFA", "ReasonID"]
act_rows = []
for n, a in enumerate(activities, start=2):
    act_rows.append([
        a["id"], a["desc"], a["area"], a["disc"], a["contractor"], a["priority"],
        a["start"], a["end"], a["planned"], a["actual"], a["progress"],
        f"=ROUND(I{n}*K{n},2)",
        f'=IF(K{n}>=1,"Done",IF(K{n}>0,"In Progress","Not Done"))',
        f'=IF(K{n}>=1,"انجام‌شده",IF(K{n}>0,"در حال انجام","انجام‌نشده"))',
        a["reason"] or None,
    ])
ws = write_sheet(
    wb, "Activities", act_headers, act_rows,
    [12, 44, 9, 13, 13, 10, 13, 13, 11, 11, 10, 11, 12, 14, 10],
    fmts={"G": "yyyy-mm-dd", "H": "yyyy-mm-dd", "I": "#,##0.00", "J": "#,##0.00",
          "K": "0%", "L": "#,##0.00"},
    fonts={c: INPUT for c in "ABCDEFGHIJKO"},
)
for row in ws.iter_rows(min_row=2):
    for c in row:
        if c.column_letter in "LMN":
            c.font = CALC

# Daily_Progress -----------------------------------------------------------
write_sheet(wb, "Daily_Progress",
            ["Date", "ActivityID", "PlannedMH_Day", "ActualMH_Day", "EarnedMH_Day"],
            daily, [13, 12, 15, 15, 15],
            fmts={"A": "yyyy-mm-dd", "C": "#,##0.00", "D": "#,##0.00", "E": "#,##0.00"},
            fonts={c: INPUT for c in "ABCDE"})

# Resources ----------------------------------------------------------------
res_rows = [[d, disc, hc, hrs, f"=C{i}*D{i}"] for i, (d, disc, hc, hrs) in enumerate(resources, start=2)]
ws = write_sheet(wb, "Resources", ["Date", "DisciplineID", "Headcount", "HoursPerPerson", "AvailableMH"],
                 res_rows, [13, 13, 12, 15, 13], fmts={"A": "yyyy-mm-dd", "E": "#,##0"},
                 fonts={c: INPUT for c in "ABCD"})
for row in ws.iter_rows(min_row=2):
    row[4].font = CALC

# Dimensions ---------------------------------------------------------------
write_sheet(wb, "Dim_Area", ["AreaID", "AreaName", "AreaNameEN"], AREAS, [10, 28, 20])
write_sheet(wb, "Dim_Discipline", ["DisciplineID", "DisciplineName", "DisciplineNameEN"], DISCIPLINES, [14, 20, 20])
write_sheet(wb, "Dim_Contractor", ["ContractorID", "ContractorName", "ContractorType"], CONTRACTORS, [14, 32, 16])
write_sheet(wb, "Dim_Reason", ["ReasonID", "ReasonName", "ReasonCategory"], [r[:3] for r in REASONS], [10, 38, 16])
write_sheet(wb, "Dim_Date",
            ["Date", "Year", "Month", "MonthName", "Day", "ShamsiDate", "ShamsiYear", "ShamsiMonth",
             "ShamsiMonthName", "ShamsiDay", "ShamsiYearMonth", "WeekdayNo", "WeekdayNameFA", "IsFriday"],
            date_rows, [13, 7, 7, 12, 6, 13, 11, 12, 16, 10, 16, 11, 14, 9],
            fmts={"A": "yyyy-mm-dd"})

OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT_XLSX)
print(f"Saved {OUT_XLSX}")

# --------------------------------------------------------------------------- expected KPIs
planned = sum(a["planned"] for a in activities)
earned = round(sum(a["earned"] for a in activities), 2)
actual = round(sum(a["actual"] for a in activities), 2)
n_done = sum(a["progress"] >= 1 for a in activities)
n_prog = sum(0 < a["progress"] < 1 for a in activities)
n_not = sum(a["progress"] == 0 for a in activities)
not_done_mh = sum(a["planned"] for a in activities if a["progress"] == 0)
available = sum(h * hrs for _, _, h, hrs in resources)

by_reason = {}
for a in activities:
    if a["reason"]:
        by_reason[a["reason"]] = by_reason.get(a["reason"], 0) + (a["planned"] - a["earned"])
reason_name = {r[0]: r[1] for r in REASONS}
reason_lines = "\n".join(
    f"| {reason_name[k]} | {v:,.2f} |" for k, v in sorted(by_reason.items(), key=lambda kv: -kv[1]))

md = f"""# Validation — expected KPI values (sample data, seed {SEED})

After building the report, the overall (unfiltered) cards must show exactly these numbers.
If they differ, check relationships and measure definitions.

| Measure | Expected value |
|---|---|
| Planned MH | {planned:,.2f} |
| Earned MH | {earned:,.2f} |
| Actual MH | {actual:,.2f} |
| Remaining MH (Planned − Earned) | {planned - earned:,.2f} |
| Completion % (Earned / Planned) | {earned / planned:.2%} |
| Productivity Index (Earned / Actual) | {earned / actual:.3f} |
| Available MH | {available:,.2f} |
| Utilization % (Actual / Available) | {actual / available:.2%} |
| Activities (total) | {len(activities)} |
| Done / In Progress / Not Done | {n_done} / {n_prog} / {n_not} |
| Not Done MH (planned MH of activities with 0%) | {not_done_mh:,.2f} |

## Unfinished MH by reason (Planned − Earned, activities with a reason)

| Reason | Unfinished MH |
|---|---|
{reason_lines}

Cross-checks inside the Excel file itself:

- Sum of `Daily_Progress[PlannedMH_Day]` = Planned MH
- Sum of `Daily_Progress[ActualMH_Day]` = Actual MH
- Sum of `Daily_Progress[EarnedMH_Day]` = Earned MH
"""
OUT_VALIDATION.parent.mkdir(parents=True, exist_ok=True)
OUT_VALIDATION.write_text(md, encoding="utf-8")
print(f"Saved {OUT_VALIDATION}")
print(f"Planned={planned:,.0f}  Earned={earned:,.2f}  Completion={earned/planned:.1%}  "
      f"Done/Prog/Not={n_done}/{n_prog}/{n_not}")
