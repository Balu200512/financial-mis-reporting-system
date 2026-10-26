\
let
    Source = Csv.Document(
        File.Contents("C:\YOUR_PROJECT_PATH\data\budgets.csv"),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders,{
        {"budget_id", type text},
        {"budget_month", type date},
        {"account_code", type text},
        {"cost_center_id", type text},
        {"budget_amount", Currency.Type},
        {"currency", type text}
    }),
    ValidRows = Table.SelectRows(ChangedTypes, each [budget_amount] >= 0)
in
    ValidRows
