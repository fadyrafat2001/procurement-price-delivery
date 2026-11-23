// UNEXECUTED M SOURCE. Create text parameter WorkbookPath with full local XLSX path.
// Create this function as fxSheet; invoke it once per data sheet below.
(sheetName as text) as table =>
let
    Source = Excel.Workbook(File.Contents(WorkbookPath), null, true),
    Data = Source{[Item=sheetName, Kind="Sheet"]}[Data],
    Headers = Table.PromoteHeaders(Data, [PromoteAllScalars=true])
in
    Headers

// Separate query named FactPO:
let
    Source = fxSheet("FactPO"),
    Typed = Table.TransformColumnTypes(Source, {{"PO_Line", type text}, {"Supplier", type text}, {"Material", type text}, {"Month", type text}, {"Ordered_kg", type number}, {"Received_kg", type number}, {"Standard_EGP_kg", type number}, {"Actual_EGP_kg", type number}, {"Delivered_Flag", type number}, {"OnTime_Flag", type number}, {"Delay_Days", type number}}, "en-US")
in
    Typed

// Separate query named DimSupplier:
let
    Source = fxSheet("DimSupplier"),
    Typed = Table.TransformColumnTypes(Source, {{"Supplier", type text}}, "en-US")
in
    Typed

// Separate query named DimMaterial:
let
    Source = fxSheet("DimMaterial"),
    Typed = Table.TransformColumnTypes(Source, {{"Material", type text}}, "en-US")
in
    Typed

// Separate query named DimMonth:
let
    Source = fxSheet("DimMonth"),
    Typed = Table.TransformColumnTypes(Source, {{"Month", type text}}, "en-US")
in
    Typed

