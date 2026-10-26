\
let
    // In Excel: Data > Get Data > From File > From Text/CSV, then open Advanced Editor.
    // Change the path below to your local project folder.
    Source = Csv.Document(
        File.Contents("C:\YOUR_PROJECT_PATH\data\transactions.csv"),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders,{
        {"transaction_id", type text},
        {"transaction_date", type date},
        {"account_code", type text},
        {"cost_center_id", type text},
        {"description", type text},
        {"reference_no", type text},
        {"amount", Currency.Type},
        {"currency", type text},
        {"status", type text},
        {"prepared_by", type text}
    }),
    TrimmedText = Table.TransformColumns(ChangedTypes,{
        {"transaction_id", Text.Trim, type text},
        {"account_code", Text.Trim, type text},
        {"cost_center_id", Text.Trim, type text},
        {"status", Text.Trim, type text}
    }),
    ValidRows = Table.SelectRows(TrimmedText, each [transaction_id] <> null and [amount] >= 0)
in
    ValidRows
