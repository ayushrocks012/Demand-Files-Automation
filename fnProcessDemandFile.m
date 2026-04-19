(FileBinary as binary, GrowthDriverName as text) =>
let
    // 1. Open the Excel workbook
    Workbook = Excel.Workbook(FileBinary, true, true),

    // 2. Ask the Vault for the valid Sheets AND valid Channels
    ValidSheetsTable = Table.SelectRows(ParameterVault[Sheets], each [Growth Driver] = GrowthDriverName),
    ValidSheetNames = List.Buffer(List.Transform(ValidSheetsTable[Sheet Name], Text.Upper)),
    
    ValidChannelsTable = Table.SelectRows(ParameterVault[Channels], each [Growth Driver] = GrowthDriverName),
    // NEW: Create an UPPERCASE join key for the Master Channels to ignore case sensitivity
    ChannelsWithUpper = Table.AddColumn(ValidChannelsTable, "JoinKey", each Text.Upper([Channel Name])),

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

        // Unpivot EVERYTHING to the right (Captures Dates AND Junk)
        Unpivoted = Table.UnpivotOtherColumns(CleanedChannels, {"Column_A", "Channel_Name"}, "Raw_Header", "Value"),
        
        // The Strict Date Parser
        ParseDates = Table.AddColumn(Unpivoted, "Parsed_Date", each 
            let 
                AsNumber = try Number.From([Raw_Header]) otherwise null,
                AsSerial = if AsNumber <> null and AsNumber > 30000 then Date.From(AsNumber) else null,
                AsText = try Date.FromText([Raw_Header]) otherwise null
            in 
                if AsSerial <> null then AsSerial else AsText, type date
        ),
        
        // Filter out all the Junk Columns
        FilteredDatesOnly = Table.SelectRows(ParseDates, each [Parsed_Date] <> null),
        
        // Format the valid dates exactly how you want them ("Jan-23")
        FormattedDates = Table.TransformColumns(FilteredDatesOnly, {{"Parsed_Date", each Date.ToText(_, "MMM-yy"), type text}}),
        
        // Cleanup the columns (Remove the raw headers, rename the clean ones)
        FinalDates = Table.RenameColumns(Table.RemoveColumns(FormattedDates, {"Raw_Header"}), {{"Parsed_Date", "Forecast_Date"}}),

        // NEW FIX: Add an UPPERCASE Join Key to the sheet data so it perfectly matches the Master Channels
        AddJoinKey = Table.AddColumn(FinalDates, "JoinKey", each Text.Upper([Channel_Name])),

        // Inner Join using the case-insensitive JoinKey!
        MergedChannels = Table.NestedJoin(AddJoinKey, {"JoinKey"}, ChannelsWithUpper, {"JoinKey"}, "ChannelMasterInfo", JoinKind.Inner),
        ExpandedType = Table.ExpandTableColumn(MergedChannels, "ChannelMasterInfo", {"Type"}, {"Data_Format_Type"}),
        
        // Clean up the temporary JoinKey and format the values
        CleanedUp = Table.RemoveColumns(ExpandedType, {"JoinKey"}),
        TypedValues = Table.TransformColumnTypes(CleanedUp, {{"Value", type number}}),
        
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
