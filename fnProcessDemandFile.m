(FileBinary as binary, GrowthDriverName as text) =>
let
    // 1. Open the Excel workbook
    Workbook = Excel.Workbook(FileBinary, true, true),

    // 2. Ask the Vault for the valid Sheets
    ValidSheetsTable = Table.SelectRows(ParameterVault[Sheets], each [Growth Driver] = GrowthDriverName),
    ValidSheetNames = List.Buffer(List.Transform(ValidSheetsTable[Sheet Name], Text.Upper)),

    // 3. Resilient Filter for Sheets
    FilteredSheets = Table.SelectRows(Workbook, each 
        List.Contains(ValidSheetNames, Text.Upper(Text.Replace([Name], "#", "."))) 
        and Record.FieldOrDefault(_, "Kind", "Sheet") = "Sheet"
    ),

    // 4. Ask the Vault where the data actually starts
    DateRowString = Table.SelectRows(ParameterVault[Config], each [KeyDetails] = "Date Row"){0}[Value],
    DateRow = Number.From(DateRowString),

    // 5. Mini-process: Just skip rows, promote headers, and tag the Brand. NO UNPIVOTING HERE.
    ProcessSheet = (SheetData as table, SheetName as text) =>
    let
        SkippedRows = Table.Skip(SheetData, DateRow - 1),
        PromotedHeaders = Table.PromoteHeaders(SkippedRows, [PromoteAllScalars=true]),
        
        ColA = try Table.ColumnNames(PromotedHeaders){0} otherwise "Column_A",
        ColB = try Table.ColumnNames(PromotedHeaders){1} otherwise "Channel_Name",
        
        RenamedHeaders = Table.RenameColumns(PromotedHeaders, {
            {ColA, "Column_A"}, 
            {ColB, "Channel_Name"}
        }),
        
        CleanSheetName = Text.Replace(SheetName, "#", "."),
        AddedBrand = Table.AddColumn(RenamedHeaders, "Brand", each CleanSheetName)
    in
        AddedBrand,

    // 6. Apply process and COMBINE all sheets in this file into one Wide Table
    ProcessedData = Table.AddColumn(FilteredSheets, "CleanData", each ProcessSheet([Data], [Name])),
    CombinedWideTable = Table.Combine(ProcessedData[CleanData])
in
    CombinedWideTable
