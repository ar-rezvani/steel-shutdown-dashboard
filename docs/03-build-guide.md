# راهنمای ساخت قدم‌به‌قدم در Power BI Desktop

1. **ساخت داده**: `python scripts/generate_sample_data.py` (یا مستقیماً از `data/ShutdownRepair_SampleData.xlsx` استفاده کنید).
2. **Power BI Desktop** را باز کنید و *Home > Transform data > Manage Parameters* را بزنید؛ پارامتر `FilePath` را مطابق `powerquery/00_FilePath_parameter.pq` بسازید و مسیر فایل اکسل را بگذارید.
3. برای هر فایل `powerquery/*.pq`: *New Source > Blank Query > Advanced Editor*، کد را paste و نام کوئری را مطابق جدول زیر بگذارید:

   | فایل | نام کوئری |
   |---|---|
   | 01_Fact_Activities.pq | Fact_Activities |
   | 02_Fact_Daily.pq | Fact_Daily |
   | 03_Fact_Resources.pq | Fact_Resources |
   | 04_Dim_Area.pq / Contractor / Discipline / Reason | Dim_Area / Dim_Contractor / Dim_Discipline / Dim_Reason |
   | 05_Dim_Date.pq | Dim_Date |

4. **Close & Apply**، سپس در *Model view* ارتباط‌ها را مطابق `docs/01-data-model.md` بسازید (احتمالاً بخشی را خودکار تشخیص می‌دهد).
5. جدول خالی `_Measures` بسازید (*Enter data*) و محتوای `dax/measures.dax` را measure به measure اضافه کنید.
6. **اعتبارسنجی**: کارت‌های بدون فیلتر را با `docs/04-validation.md` مقایسه کنید.
7. تم را بارگذاری کنید: *View > Themes > Browse for themes > theme/steel-theme.json*.
8. صفحات را مطابق `docs/02-dashboard-design.md` بسازید.
9. ذخیره: برای کار تیمی با Git، *File > Save as* را با فرمت **Power BI Project (.pbip)** انجام دهید (فایل‌های متنی، قابل diff). در غیر این صورت `.pbix` را با Git LFS نگه دارید (`.gitattributes` آماده است).

## به‌روزرسانی داده
فایل اکسل را جایگزین/ویرایش کنید، سپس **Refresh**. نام شیت‌ها و ستون‌ها نباید تغییر کند.

## اتصال به داده واقعی کارفرما
فقط کوئری‌های Power Query را تغییر دهید (منبع: SAP/CMMS/SQL)؛ مدل و DAX ثابت می‌مانند.
