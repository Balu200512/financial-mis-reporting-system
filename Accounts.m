\
let
    Source = Csv.Document(
        File.Contents("C:\YOUR_PROJECT_PATH\data\accounts.csv"),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders,{
        {"account_code", type text},
        {"account_name", type text},
        {"account_type", type text}
    }),
    Trimmed = Table.TransformColumns(ChangedTypes,{
        {"account_code", Text.Trim, type text},
        {"account_name", Text.Trim, type text},
        {"account_type", Text.Trim, type text}
    })
in
    Trimmed
