# طراحی صفحات داشبورد

اندازه صفحه: 16:9 (1280×720). تم: `theme/steel-theme.json`. فونت: Tahoma. چیدمان راست‌به‌چپ (دستی).

## صفحه ۱ — خلاصه مدیریتی
- **کارت‌ها**: Planned MH, Earned MH, Remaining MH, Completion %, Not Done Count, Productivity Index
- **Gauge/KPI**: Completion % (رنگ شرطی با `Completion Color`)
- **Donut**: تعداد فعالیت‌ها بر اساس Status
- **Clustered bar**: Planned MH در برابر Earned MH بر اساس تخصص
- **Slicer**: ناحیه، تخصص، پیمانکار، اولویت

## صفحه ۲ — تحلیل بر اساس تخصص / ناحیه / پیمانکار
- Matrix: ردیف=تخصص، ستون‌ها: Planned, Earned, Remaining, Completion %, Productivity (با Data bars)
- Bar chart: Completion % بر اساس ناحیه
- Bar chart: Remaining MH بر اساس پیمانکار

## صفحه ۳ — تحلیل علل عدم انجام (Pareto)
- **Line and clustered column**: محور = `Dim_Reason[ReasonName]` (مرتب نزولی بر اساس Unfinished MH)، ستون = `Unfinished MH`، خط = `Pareto Cumulative %`
- فیلتر سطح visual: `Unfinished MH` is not blank
- Treemap: `ReasonCategory` × Unfinished MH
- Table drill-through: فعالیت‌های متأثر از هر علت
- زیرعنوان پویا: measure `Reasons Subtitle`

## صفحه ۴ — جدول ریز فعالیت‌ها
- Table: ActivityID, Description, Area, Discipline, Contractor, Priority, Planned MH, Earned MH, Progress (Data bars), StatusFA, ReasonName
- فیلتر روی وضعیت و علت؛ **Drill-through** از صفحه ۳ به اینجا

## صفحه ۵ — روند و منابع
- Line chart (منحنی S): محور `Dim_Date[ShamsiDate]`، مقادیر: Cumulative Planned / Earned / Actual
- Column + Line: Available MH در برابر Actual MH (Daily) و Utilization % به تفکیک روز
- Matrix: تخصص × روز با Utilization % (Conditional formatting)
- کارت: SPI

## نکات UX
- Bookmark برای «فقط انجام‌نشده‌ها» و «فقط فعالیت‌های Critical»
- Tooltip page برای نمایش جزئیات فعالیت
- Utilization فقط به تفکیک **تخصص** یا **روز** معنی دارد (نه ناحیه/پیمانکار).
