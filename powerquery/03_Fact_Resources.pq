let
    Source   = Excel.Workbook(File.Contents(FilePath), null, true),
    Sheet    = Source{[Item = "Resources", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Typed    = Table.TransformColumnTypes(Promoted, {
        {"Date", type date}, {"DisciplineID", type text},
        {"Headcount", Int64.Type}, {"HoursPerPerson", type number}, {"AvailableMH", type number}
    }, "en-US"),
    NonEmpty = Table.SelectRows(Typed, each [Date] <> null)
in
    NonEmpty
