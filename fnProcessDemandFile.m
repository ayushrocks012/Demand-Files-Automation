(FileBinary as binary, GrowthDriverName as text) =>
let
    // 1. Open Workbook
    Workbook = Excel.Workbook(FileBinary, true, true),

    // 2. Filter for valid sheets based on your Sheet Master
    ValidSheetsTable = Table.SelectRows(ParameterVault[Sheets], each [Growth Driver] = GrowthDriverName),
    ValidSheetNames = List.Buffer(List.Transform(ValidSheetsTable[Sheet Name], Text.Upper)),
    FilteredSheets = Table.SelectRows(Workbook, each 
        List.Contains(ValidSheetNames, Text.Upper(Text.Replace([Name], "#", "."))) 
        and Record.FieldOrDefault(_, "Kind", "Sheet") = "Sheet"
    ),

    // 3. YOUR EXACT RECORDED LOGIC APPLIED TO EACH SHEET
    ProcessSheet = (SheetData as table, SheetName as text) =>
    let
        // #"Removed Columns" = Table.RemoveColumns(...,{"Column1"})
        // (Wrapped in a 'try' just in case a user accidentally deletes Column A before saving)
        RemovedCol1 = try Table.RemoveColumns(SheetData, {"Column1"}) otherwise SheetData,
        
        // #"Removed Top Rows" = Table.Skip(#"Removed Columns",65)
        SkippedRows = Table.Skip(RemovedCol1, 65),
        
        // #"Promoted Headers" = Table.PromoteHeaders(#"Removed Blank Rows", [PromoteAllScalars=true])
        PromotedHeaders = Table.PromoteHeaders(SkippedRows, [PromoteAllScalars=true]),
        
        // Standardize the first column name to "Channel_Name" so the Master Query can join on it later
        FirstColName = Table.ColumnNames(PromotedHeaders){0},
        RenamedLabel = Table.RenameColumns(PromotedHeaders, {{FirstColName, "Channel_Name"}}),
        
        // #"Filtered Rows1" = Table.SelectRows(...) -> Your step to filter out null/blank rows
        RemoveBlanks = Table.SelectRows(RenamedLabel, each [Channel_Name] <> null and [Channel_Name] <> ""),
        
        // Tag it with the Brand name
        CleanSheetName = Text.Replace(SheetName, "#", "."),
        AddedBrand = Table.AddColumn(RemoveBlanks, "Brand", each CleanSheetName)
    in
        AddedBrand,

    // 4. Run your logic on every sheet and stack them into one Wide Table
    ProcessedData = Table.AddColumn(FilteredSheets, "CleanData", each ProcessSheet([Data], [Name])),
    CombinedWideTable = Table.Combine(ProcessedData[CleanData])
in
    CombinedWideTable
