let
    Source   = Excel.Workbook(File.Contents(FilePath), null, true),
    Sheet    = Source{[Item = "Dim_Discipline", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    AllText  = Table.TransformColumnTypes(Promoted, List.Transform(Table.ColumnNames(Promoted), each {_, type text}))
in
    AllText
