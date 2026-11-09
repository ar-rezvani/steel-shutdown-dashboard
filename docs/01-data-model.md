# مدل داده (Star Schema)

## جدول‌ها

| جدول | نوع | شیت اکسل | شرح |
|---|---|---|---|
| `Fact_Activities` | Fact | Activities | هر ردیف یک فعالیت تعمیراتی: نفر‑ساعت برنامه‌ای/واقعی/کسب‌شده، درصد پیشرفت، وضعیت، علت عدم انجام |
| `Fact_Daily` | Fact | Daily_Progress | ریز روزانه هر فعالیت (برای منحنی S) |
| `Fact_Resources` | Fact | Resources | نفر‑ساعت در اختیار هر تخصص در هر روز |
| `Dim_Area` | Dim | Dim_Area | ناحیه / واحد (کوره، ریخته‌گری، ...) |
| `Dim_Discipline` | Dim | Dim_Discipline | تخصص (مکانیک، برق، ...) |
| `Dim_Contractor` | Dim | Dim_Contractor | پیمانکار / تیم داخلی |
| `Dim_Reason` | Dim | Dim_Reason | علل عدم انجام + دسته‌بندی |
| `Dim_Date` | Dim | Dim_Date | تقویم میلادی + شمسی |

## ارتباط‌ها (همه Many-to-One، جهت فیلتر Single از Dim به Fact)

| از (Many) | به (One) |
|---|---|
| `Fact_Activities[AreaID]` | `Dim_Area[AreaID]` |
| `Fact_Activities[DisciplineID]` | `Dim_Discipline[DisciplineID]` |
| `Fact_Activities[ContractorID]` | `Dim_Contractor[ContractorID]` |
| `Fact_Activities[ReasonID]` | `Dim_Reason[ReasonID]` |
| `Fact_Daily[ActivityID]` | `Fact_Activities[ActivityID]` |
| `Fact_Daily[Date]` | `Dim_Date[Date]` |
| `Fact_Resources[DisciplineID]` | `Dim_Discipline[DisciplineID]` |
| `Fact_Resources[Date]` | `Dim_Date[Date]` |

> `Fact_Daily` از طریق `Fact_Activities` به ناحیه/تخصص/پیمانکار فیلتر می‌شود.

## تنظیمات مدل
- `Dim_Date[Date]` را **Mark as date table** کنید.
- `Dim_Date[ShamsiMonthName]` را بر اساس `ShamsiMonth` و `WeekdayNameFA` را بر اساس `WeekdayNo` **Sort by column** کنید.
- ستون‌های ID را در حالت **Hidden** بگذارید و فقط نام‌ها را نمایش دهید.
- `Fact_Activities[Progress]` فرمت درصد؛ `Status` با ترتیب دستی: Not Done, In Progress, Done (یا جدول `Dim_Status` اضافه کنید).

## تعریف شاخص‌ها

| شاخص | تعریف |
|---|---|
| Planned MH | نفر‑ساعت برنامه‌ریزی‌شده |
| Earned MH | Planned × Progress (ساعت «کسب‌شده») |
| Actual MH | ساعت واقعاً مصرف‌شده |
| Completion % | Earned / Planned |
| Productivity Index | Earned / Actual (بزرگ‌تر از ۱ یعنی بهره‌وری بالاتر از برنامه) |
| Unfinished MH | Planned − Earned |
| Available MH | Headcount × Hours (نفر‑ساعت در اختیار) |
| Utilization % | Actual / Available |
| SPI | Cumulative Earned / Cumulative Planned |

> این تعاریف باید قبل از تحویل با کارفرما تأیید شوند (به‌خصوص «انجام‌شده» = Earned یا Actual).
