(FileBinary as binary, GrowthDriverName as text) =>
let
    // 1. Open the Excel workbook
    Workbook = Excel.Workbook(FileBinary, true, true),

    // 2. Ask the Vault which sheets belong to this specific Growth Driver
    ValidSheetsTable = Table.SelectRows(ParameterVault[Sheets], each [Growth Driver] = GrowthDriverName),
    ValidSheetNames = List.Buffer(ValidSheetsTable[Sheet Name]),

    // 3. Filter the workbook to ONLY include those valid sheets
    FilteredSheets = Table.SelectRows(Workbook, each List.Contains(ValidSheetNames, [Item]) and [Kind] = "Sheet"),

    // 4. Ask the Vault where the data actually starts (e.g., Row 66)
    DateRowString = Table.SelectRows(ParameterVault[Config], each [KeyDetails] = "Date Row"){0}[Value],
    DateRow = Number.From(DateRowString),

    // 5. Build the mini-process to clean a SINGLE sheet
    ProcessSheet = (SheetData as table, SheetName as text) =>
    let
        // Skip the junk rows above the header (e.g., skip 65 rows if data starts at 66)
        SkippedRows = Table.Skip(SheetData, DateRow - 1),
        // Promote the new top row to be the column headers
        PromotedHeaders = Table.PromoteHeaders(SkippedRows, [PromoteAllScalars=true]),
        
        // Dynamically grab the names of the first two columns to standardize them
        ColAName = Table.ColumnNames(PromotedHeaders){0},
        LabelColumnName = Table.ColumnNames(PromotedHeaders){1},
        
        // Rename them to safe, standardized names so the code never breaks
        StandardizedHeaders = Table.RenameColumns(PromotedHeaders, {
            {ColAName, "Column_A"}, 
            {LabelColumnName, "Channel_Name"}
        }),
        
        // Remove any row where the Channel Name is completely blank
        FilteredBlanks = Table.SelectRows(StandardizedHeaders, each [Channel_Name] <> null and [Channel_Name] <> ""),

        // Unpivot all columns EXCEPT Column_A and Channel_Name (meaning, unpivot the dates)
        AllStandardColumns = Table.ColumnNames(FilteredBlanks),
        ColumnsToUnpivot = List.Skip(AllStandardColumns, 2),
        Unpivoted = Table.Unpivot(FilteredBlanks, ColumnsToUnpivot, "Forecast_Date", "Value"),
        
        // Add a column with the Brand (Sheet) Name
        AddedBrand = Table.AddColumn(Unpivoted, "Brand", each SheetName)
    in
        AddedBrand,

    // 6. Apply this mini-process to every valid sheet in the workbook
    ProcessedData = Table.AddColumn(FilteredSheets, "CleanData", each ProcessSheet([Data], [Item])),
    
    // 7. Expand the cleaned data into one massive flat table for this file
    ExpandedData = Table.ExpandTableColumn(ProcessedData, "CleanData", 
        {"Brand", "Column_A", "Channel_Name", "Forecast_Date", "Value"}
    ),

    // 8. Clean up columns we no longer need (like the raw binary sheet data)
    FinalTable = Table.SelectColumns(ExpandedData, {"Brand", "Column_A", "Channel_Name", "Forecast_Date", "Value"})
in
    FinalTable
