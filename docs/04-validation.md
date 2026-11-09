# Validation — expected KPI values (sample data, seed 42)

After building the report, the overall (unfiltered) cards must show exactly these numbers.
If they differ, check relationships and measure definitions.

| Measure | Expected value |
|---|---|
| Planned MH | 22,280.00 |
| Earned MH | 16,634.45 |
| Actual MH | 18,267.63 |
| Remaining MH (Planned − Earned) | 5,645.55 |
| Completion % (Earned / Planned) | 74.66% |
| Productivity Index (Earned / Actual) | 0.911 |
| Available MH | 25,930.00 |
| Utilization % (Actual / Available) | 70.45% |
| Activities (total) | 120 |
| Done / In Progress / Not Done | 72 / 25 / 23 |
| Not Done MH (planned MH of activities with 0%) | 3,720.00 |

## Unfinished MH by reason (Planned − Earned, activities with a reason)

| Reason | Unfinished MH |
|---|---|
| کمبود قطعه یدکی / مواد | 1,537.70 |
| تأخیر پیمانکار | 1,150.90 |
| تأخیر صدور مجوز کار | 883.70 |
| کمبود نیروی انسانی | 735.70 |
| محدودیت زمان توقف | 605.00 |
| کشف کار اضافه (Extra Scope) | 282.90 |
| مشکل دسترسی / داربست | 264.65 |
| خرابی تجهیزات کمکی (جرثقیل/ابزار) | 185.00 |

Cross-checks inside the Excel file itself:

- Sum of `Daily_Progress[PlannedMH_Day]` = Planned MH
- Sum of `Daily_Progress[ActualMH_Day]` = Actual MH
- Sum of `Daily_Progress[EarnedMH_Day]` = Earned MH
