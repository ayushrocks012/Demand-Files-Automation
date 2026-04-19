let
    // 1. Get the list of Target Files from our Vault
    TargetFiles = ParameterVault[Files],
    Source = SharePoint.Files("https://abbott.sharepoint.com/sites/GB-AN-HeadOffice", [ApiVersion = 15]),
    
    // 2. Match the exact File Names (This replaces your hardcoded folder path with dynamic matching)
    MergedFiles = Table.NestedJoin(TargetFiles, {"File Name"}, Source, {"Name"}, "SP_Data", JoinKind.Inner),
    ExpandedSP = Table.ExpandTableColumn(MergedFiles, "SP_Data", {"Content"}, {"FileBinary"}),
    
    // 3. Extract Metadata from File Names
    NormalizeName = Table.TransformColumns(ExpandedSP, {{"File Name", each Text.BeforeDelimiter(Text.Replace(_, "-", "_"), "."), type text}}),
    SplitName = Table.AddColumn(NormalizeName, "NameParts", each Text.Split([File Name], "_")),
    ExtractMeta = Table.TransformRows(SplitName, each _ & [
        Affiliate = [NameParts]{0}?, Growth_Driver_Name = [NameParts]{1}?, File_Type = [NameParts]{2}?,
        Demand_Plan_Month = [NameParts]{3}?, Actuals_Month = [NameParts]{4}?, Version = Text.Combine(List.Skip([NameParts], 5), "_")
    ]),
    MetaTable = Table.FromRecords(ExtractMeta),
    
    // 4. Run Your Engine (Returns the stacked Wide Tables)
    InvokeEngine = Table.AddColumn(MetaTable, "CleanData", each fnProcessDemandFile([FileBinary], [Growth_Driver_Name])),
    
    // 5. Inject Metadata into the Wide Tables
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
    
    // 6. Stack everything into ONE massive wide table
    MassiveWideTable = Table.Combine(InjectMeta[WideDataWithMeta]),

    // 7. Filter using your Channel Master (Growth Driver + Channel Name)
    ValidChannels = Table.AddColumn(ParameterVault[Channels], "JoinKey", each try Text.Upper([Growth Driver] & "|" & Text.Trim([Channel Name])) otherwise ""),
    WideWithJoinKey = Table.AddColumn(MassiveWideTable, "JoinKey", each try Text.Upper([Growth_Driver_Name] & "|" & Text.Trim([Channel_Name])) otherwise ""),
    MergedChannels = Table.NestedJoin(WideWithJoinKey, {"JoinKey"}, ValidChannels, {"JoinKey"}, "ChannelMasterInfo", JoinKind.Inner),
    FilteredWideTable = Table.ExpandTableColumn(MergedChannels, "ChannelMasterInfo", {"Type"}, {"Data_Format_Type"}),

    // 8. Unpivot ONLY the Dates
    MetaCols = {"Year Folder", "Month Folder", "Affiliate", "Growth_Driver_Name", "Brand", "File_Type", "Demand_Plan_Month", "Actuals_Month", "Version", "Channel_Name", "Data_Format_Type", "JoinKey"},
    Unpivoted = Table.UnpivotOtherColumns(FilteredWideTable, MetaCols, "Raw_Date", "Value"),

    // 9. Clean Dates & Drop Ghost Columns
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

    // 10. Final Cleanup
    TypedValues = Table.TransformColumnTypes(FilterValidDates, {{"Value", type number}}),
    FinalTable = Table.SelectColumns(TypedValues, {
        "Year Folder", "Month Folder", "Affiliate", "Growth_Driver_Name", "Brand", 
        "File_Type", "Demand_Plan_Month", "Actuals_Month", "Version", 
        "Channel_Name", "Data_Format_Type", "Forecast_Date", "Value"
    })
in
    FinalTable
