(FileBinary as binary, GrowthDriverName as text) =>
let
    // 1. Open the Excel workbook
    Workbook = Excel.Workbook(FileBinary, true, true),

    // 2. Ask the Vault for the valid Sheets AND valid Channels
    ValidSheetsTable = Table.SelectRows(ParameterVault[Sheets], each [Growth Driver] = GrowthDriverName),
    ValidSheetNames = List.Buffer(List.Transform(ValidSheetsTable[Sheet Name], Text.Upper)),
    ValidChannelsTable = Table.SelectRows(ParameterVault[Channels], each [Growth Driver] = GrowthDriverName),

    // 3. Resilient Filter for Sheets
    FilteredSheets = Table.SelectRows(Workbook, each 
        List.Contains(ValidSheetNames, Text.Upper(Text.Replace([Name], "#", "."))) 
        and Record.FieldOrDefault(_, "Kind", "Sheet") = "Sheet"
    ),

    // 4. Ask the Vault where the data actually starts
    DateRowString = Table.SelectRows(ParameterVault[Config], each [KeyDetails] = "Date Row"){0}[Value],
    DateRow = Number.From(DateRowString),

    // 5. Build the mini-process to clean a SINGLE sheet
    ProcessSheet = (SheetData as table, SheetName as text) =>
    let
        SkippedRows = Table.Skip(SheetData, DateRow - 1),
        PromotedHeaders = Table.PromoteHeaders(SkippedRows, [PromoteAllScalars=true]),
        
        ColAName = Table.ColumnNames(PromotedHeaders){0},
        LabelColumnName = Table.ColumnNames(PromotedHeaders){1},
        
        StandardizedHeaders = Table.RenameColumns(PromotedHeaders, {
            {ColAName, "Column_A"}, 
            {LabelColumnName, "Channel_Name"}
        }),
        
        CleanedChannels = Table.TransformColumns(StandardizedHeaders, {{"Channel_Name", each try Text.Trim(_) otherwise _}}),

        // NEW FIX 1: Unpivot FIRST. "UnpivotOtherColumns" safely grabs all date columns (Jan-23 to Dec-27+)
        Unpivoted = Table.UnpivotOtherColumns(CleanedChannels, {"Column_A", "Channel_Name"}, "Forecast_Date", "Value"),
        
        // NEW FIX 2: Do the Inner Join AFTER the dates are flattened.
        MergedChannels = Table.NestedJoin(Unpivoted, {"Channel_Name"}, ValidChannelsTable, {"Channel Name"}, "ChannelMasterInfo", JoinKind.Inner),
        ExpandedType = Table.ExpandTableColumn(MergedChannels, "ChannelMasterInfo", {"Type"}, {"Data_Format_Type"}),

        // NEW FIX 3: Date Cleanup. If Excel turned "Jan-23" into "44927" during headers, change it back.
        CleanDates = Table.TransformColumns(ExpandedType, {
            {"Forecast_Date", each 
                if try Number.From(_) > 30000 otherwise false 
                then Date.ToText(Date.From(Number.From(_)), "MMM-yy") 
                else _ 
            }
        }),

        // Ensure the Value column is strictly treated as a number
        TypedValues = Table.TransformColumnTypes(CleanDates, {{"Value", type number}}),
        
        CleanSheetName = Text.Replace(SheetName, "#", "."),
        AddedBrand = Table.AddColumn(TypedValues, "Brand", each CleanSheetName)
    in
        AddedBrand,

    // 6. Apply this mini-process to every valid sheet
    ProcessedData = Table.AddColumn(FilteredSheets, "CleanData", each ProcessSheet([Data], [Name])),
    
    // 7. Expand the cleaned data
    ExpandedData = Table.ExpandTableColumn(ProcessedData, "CleanData", 
        {"Brand", "Channel_Name", "Data_Format_Type", "Forecast_Date", "Value"}
    ),

    // 8. Select the final columns
    FinalTable = Table.SelectColumns(ExpandedData, {"Brand", "Channel_Name", "Data_Format_Type", "Forecast_Date", "Value"})
in
    FinalTable
