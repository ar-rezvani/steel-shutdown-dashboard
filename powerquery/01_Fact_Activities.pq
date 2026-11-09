let
    Source   = Excel.Workbook(File.Contents(FilePath), null, true),
    Sheet    = Source{[Item = "Activities", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Typed    = Table.TransformColumnTypes(Promoted, {
        {"ActivityID", type text}, {"Description", type text}, {"AreaID", type text},
        {"DisciplineID", type text}, {"ContractorID", type text}, {"Priority", type text},
        {"PlannedStart", type date}, {"PlannedEnd", type date},
        {"PlannedMH", type number}, {"ActualMH", type number}, {"Progress", Percentage.Type},
        {"EarnedMH", type number}, {"Status", type text}, {"StatusFA", type text},
        {"ReasonID", type text}
    }, "en-US"),
    // keep only real activity rows
    NonEmpty = Table.SelectRows(Typed, each [ActivityID] <> null and [ActivityID] <> ""),
    // unfinished hours (used by the "reason" analysis)
    AddUnfinished = Table.AddColumn(NonEmpty, "UnfinishedMH", each [PlannedMH] - [EarnedMH], type number)
in
    AddUnfinished
