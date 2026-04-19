(FileBinary as binary, GrowthDriverName as text) =>
let
    // 1. Open Workbook (Your #"Imported Excel Workbook" step)
    Workbook = Excel.Workbook(FileBinary, true, true),

    // 2. Filter for valid sheets based on your Sheet Master
    ValidSheetsTable = Table.SelectRows(ParameterVault[Sheets], each [Growth Driver] = GrowthDriverName),
    ValidSheetNames = List.Buffer(List.Transform(ValidSheetsTable[Sheet Name], Text.Upper)),
    FilteredSheets = Table.SelectRows(Workbook, each 
        List.Contains(ValidSheetNames, Text.Upper(Text.Replace([Name], "#", "."))) 
        and Record.FieldOrDefault(_, "Kind", "Sheet") = "Sheet"
    ),

    // 3. Get the dynamic row number (66)
    DateRowString = Table.SelectRows(ParameterVault[Config], each [KeyDetails] = "Date Row"){0}[Value],
    DateRow = Number.From(DateRowString),

    // 4. YOUR EXACT LOGIC APPLIED TO EACH SHEET
    ProcessSheet = (SheetData as table, SheetName as text) =>
    let
        // Your #"Removed Columns" step
        RemovedCol1 = try Table.RemoveColumns(SheetData, {"Column1"}) otherwise SheetData,
        
        // Your #"Removed Top Rows" step (65)
        SkippedRows = Table.Skip(RemovedCol1, DateRow - 1),
        
        // Your #"Promoted Headers" step
        PromotedHeaders = Table.PromoteHeaders(SkippedRows, [PromoteAllScalars=true]),
        
        // Find the label column (which is now the first column) and rename it for consistency
        FirstColName = Table.ColumnNames(PromotedHeaders){0},
        RenamedLabel = Table.RenameColumns(PromotedHeaders, {{FirstColName, "Channel_Name"}}),
        
        // Your step to filter out null/blank rows
        RemoveBlanks = Table.SelectRows(RenamedLabel, each [Channel_Name] <> null and [Channel_Name] <> ""),
        
        // Tag it with the Brand name
        CleanSheetName = Text.Replace(SheetName, "#", "."),
        AddedBrand = Table.AddColumn(RemoveBlanks, "Brand", each CleanSheetName)
    in
        AddedBrand,

    // 5. Run your logic on every sheet and stack them into one Wide Table
    ProcessedData = Table.AddColumn(FilteredSheets, "CleanData", each ProcessSheet([Data], [Name])),
    CombinedWideTable = Table.Combine(ProcessedData[CleanData])
in
    CombinedWideTable
