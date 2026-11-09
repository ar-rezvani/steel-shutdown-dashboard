let
    Source   = Excel.Workbook(File.Contents(FilePath), null, true),
    Sheet    = Source{[Item = "Dim_Date", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Typed    = Table.TransformColumnTypes(Promoted, {
        {"Date", type date}, {"Year", Int64.Type}, {"Month", Int64.Type}, {"MonthName", type text},
        {"Day", Int64.Type}, {"ShamsiDate", type text}, {"ShamsiYear", Int64.Type},
        {"ShamsiMonth", Int64.Type}, {"ShamsiMonthName", type text}, {"ShamsiDay", Int64.Type},
        {"ShamsiYearMonth", type text}, {"WeekdayNo", Int64.Type}, {"WeekdayNameFA", type text},
        {"IsFriday", Int64.Type}
    }, "en-US")
in
    Typed
