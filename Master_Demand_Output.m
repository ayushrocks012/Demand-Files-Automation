let
    // 1 to 5: Standard Extraction & Metadata Setup
    TargetFiles = ParameterVault[Files],
    Source = SharePoint.Files("https://abbott.sharepoint.com/sites/GB-AN-HeadOffice", [ApiVersion = 15]),
    
    MergedFiles = Table.NestedJoin(TargetFiles, {"File Name"}, Source, {"Name"}, "SP_Data", JoinKind.Inner),
    ExpandedSP = Table.ExpandTableColumn(MergedFiles, "SP_Data", {"Content"}, {"FileBinary"}),
    
    NormalizeName = Table.TransformColumns(ExpandedSP, {
        {"File Name", each Text.BeforeDelimiter(Text.Replace(_, "-", "_"), "."), type text}
    }),
    
    SplitName = Table.AddColumn(NormalizeName, "NameParts", each Text.Split([File Name], "_")),
    ExtractMeta = Table.TransformRows(SplitName, each _ & [
        Affiliate = [NameParts]{0}?,
        Growth_Driver_Name = [NameParts]{1}?,
        File_Type = [NameParts]{2}?,
        Demand_Plan_Month = [NameParts]{3}?,
        Actuals_Month = [NameParts]{4}?,
        Version = Text.Combine(List.Skip([NameParts], 5), "_")
    ]),
    MetaTable = Table.FromRecords(ExtractMeta),
    
    // 6. Invoke Engine (This returns the Wide Tables)
    InvokeEngine = Table.AddColumn(MetaTable, "CleanData", each fnProcessDemandFile([FileBinary], [Growth_Driver_Name])),
    
    // 7. Inject Metadata into the Wide Tables BEFORE appending them
    InjectMeta = Table.AddColumn(InvokeEngine, "WideDataWithMeta", each 
        let 
            tbl = [CleanData], Rec = _,
            t1 = Table.AddColumn(tbl, "Affiliate", (x) => Rec[Affiliate]),
            t2 = Table.AddColumn(t1, "Growth_Driver_Name", (x) => Rec[Growth_Driver_Name]),
            t3 = Table.AddColumn(t2, "File_Type", (x) => Rec[File_Type]),
            t4 = Table.AddColumn(t3, "Demand_Plan_Month", (x) => Rec[Demand_Plan_Month]),
            t5 = Table.AddColumn(t4, "Actuals_Month", (x) => Rec[Actuals_Month]),
            t6 = Table.AddColumn(t5, "Version", (x) => Rec[Version]),
            t7 = Table.AddColumn(t6, "Year Folder", (x) => Rec[#"Year Folder"]),
            t8 = Table.AddColumn(t7, "Month Folder", (x) => Rec[#"Month Folder"])
        in t8
    ),
    
    // 8. CREATE THE MASSIVE WIDE TABLE (Stacks every file together automatically)
    MassiveWideTable = Table.Combine(InjectMeta[WideDataWithMeta]),

    // 9. YOUR LOGIC: Filter using the Growth Driver + Channel Name combination
    ValidChannels = Table.AddColumn(ParameterVault[Channels], "JoinKey", each try Text.Upper([Growth Driver] & "|" & Text.Trim([Channel Name])) otherwise ""),
    WideWithJoinKey = Table.AddColumn(MassiveWideTable, "JoinKey", each try Text.Upper([Growth_Driver_Name] & "|" & Text.Trim([Channel_Name])) otherwise ""),
    
    MergedChannels = Table.NestedJoin(WideWithJoinKey, {"JoinKey"}, ValidChannels, {"JoinKey"}, "ChannelMasterInfo", JoinKind.Inner),
    FilteredWideTable = Table.ExpandTableColumn(MergedChannels, "ChannelMasterInfo", {"Type"}, {"Data_Format_Type"}),

    // 10. UNPIVOT ONLY THE DATES
    MetaCols = {"Year Folder", "Month Folder", "Affiliate", "Growth_Driver_Name", "Brand", "File_Type", "Demand_Plan_Month", "Actuals_Month", "Version", "Column_A", "Channel_Name", "Data_Format_Type", "JoinKey"},
    Unpivoted = Table.UnpivotOtherColumns(FilteredWideTable, MetaCols, "Raw_Date", "Value"),

    // 11. CLEAN DATES & DROP GHOST COLUMNS (e.g., drops "0.00%_1")
    CleanDates = Table.AddColumn(Unpivoted, "Forecast_Date", each 
        let 
            AsNum = try Number.From([Raw_Date]) otherwise null,
            AsSer = if AsNum <> null and AsNum > 30000 then Date.From(AsNum) else null,
            AsTxt = try Date.FromText([Raw_Date]) otherwise null,
            ValidDate = if AsSer <> null then AsSer else AsTxt
        in 
            if ValidDate <> null then Date.ToText(ValidDate, "MMM-yy") else null
    ),
    FilterValidDates = Table.SelectRows(CleanDates, each [Forecast_Date] <> null),

    // 12. FINAL CLEANUP
    TypedValues = Table.TransformColumnTypes(FilterValidDates, {{"Value", type number}}),
    FinalTable = Table.SelectColumns(TypedValues, {
        "Year Folder", "Month Folder", "Affiliate", "Growth_Driver_Name", "Brand", 
        "File_Type", "Demand_Plan_Month", "Actuals_Month", "Version", 
        "Channel_Name", "Data_Format_Type", "Forecast_Date", "Value"
    })
in
    FinalTable
