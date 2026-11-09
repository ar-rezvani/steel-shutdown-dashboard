let
    Source   = Excel.Workbook(File.Contents(FilePath), null, true),
    Sheet    = Source{[Item = "Daily_Progress", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Typed    = Table.TransformColumnTypes(Promoted, {
        {"Date", type date}, {"ActivityID", type text},
        {"PlannedMH_Day", type number}, {"ActualMH_Day", type number}, {"EarnedMH_Day", type number}
    }, "en-US"),
    NonEmpty = Table.SelectRows(Typed, each [Date] <> null and [ActivityID] <> null)
in
    NonEmpty
