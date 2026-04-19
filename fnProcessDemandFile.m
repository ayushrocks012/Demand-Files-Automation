(FileBinary as binary, GrowthDriverName as text) =>
let
    // 1. Open the Excel workbook
    Workbook = Excel.Workbook(FileBinary, true, true),

    // 2. Ask the Vault which sheets belong to this specific Growth Driver
    ValidSheetsTable = Table.SelectRows(ParameterVault[Sheets], each [Growth Driver] = GrowthDriverName),
    ValidSheetNames = List.Buffer(List.Transform(ValidSheetsTable[Sheet Name], Text.Upper)),

    // 3. Resilient Filter: Handles missing columns, replaces "#", and ignores Case Sensitivity
    FilteredSheets = Table.SelectRows(Workbook, each 
        List.Contains(ValidSheetNames, Text.Upper(Text.Replace([Name], "#", "."))) 
        and Record.FieldOrDefault(_, "Kind", "Sheet") = "Sheet"
    ),

    // 4. Ask the Vault where the data actually starts (e.g., Row 66)
    DateRowString = Table.SelectRows(ParameterVault[Config], each [KeyDetails] = "Date Row"){0}[Value],
    DateRow = Number.From(DateRowString),

    // 5. Build the mini-process to clean a SINGLE sheet
    ProcessSheet = (SheetData as table, SheetName as text) =>
    let
        // Skip the junk rows above the header
        SkippedRows = Table.Skip(SheetData, DateRow - 1),
        // Promote the new top row to be the column headers
        PromotedHeaders = Table.PromoteHeaders(SkippedRows, [PromoteAllScalars=true]),
        
        // Dynamically grab the names of the first two columns to standardize them
        ColAName = Table.ColumnNames(PromotedHeaders){0},
        LabelColumnName = Table.ColumnNames(PromotedHeaders){1},
        
        // Rename them to safe, standardized names
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
        
        // Convert the hash back to a dot for your final, clean Output file
        CleanSheetName = Text.Replace(SheetName, "#", "."),
        AddedBrand = Table.AddColumn(Unpivoted, "Brand", each CleanSheetName)
    in
        AddedBrand,

    // 6. Apply this mini-process to every valid sheet
    ProcessedData = Table.AddColumn(FilteredSheets, "CleanData", each ProcessSheet([Data], [Name])),
    
    // 7. Expand the cleaned data (REMOVED: Column_A)
    ExpandedData = Table.ExpandTableColumn(ProcessedData, "CleanData", 
        {"Brand", "Channel_Name", "Forecast_Date", "Value"}
    ),

    // 8. Clean up columns we no longer need (REMOVED: Column_A)
    FinalTable = Table.SelectColumns(ExpandedData, {"Brand", "Channel_Name", "Forecast_Date", "Value"})
in
    FinalTable
