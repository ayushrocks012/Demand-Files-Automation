let
    // 1. Get the list of Target Files from our Vault
    TargetFiles = ParameterVault[Files],
    
    // 2. Get the master list of all files in SharePoint
    Source = SharePoint.Files("https://abbott.sharepoint.com/sites/GB-AN-HeadOffice", [ApiVersion = 15]),
    
    // 3. Match the exact File Names
    MergedFiles = Table.NestedJoin(TargetFiles, {"File Name"}, Source, {"Name"}, "SP_Data", JoinKind.Inner),
    ExpandedSP = Table.ExpandTableColumn(MergedFiles, "SP_Data", {"Content"}, {"FileBinary"}),
    
    // 4. Normalize the File Name 
    NormalizeName = Table.TransformColumns(ExpandedSP, {
        {"File Name", each Text.BeforeDelimiter(Text.Replace(_, "-", "_"), "."), type text}
    }),
    
    // 5. Extract Metadata 
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
    
    // 6. Invoke the Engine
    InvokeEngine = Table.AddColumn(MetaTable, "CleanData", each fnProcessDemandFile([FileBinary], [Growth_Driver_Name])),
    
    // 7. Expand the final clean data from the engine (NEW: Included Data_Format_Type)
    ExpandedData = Table.ExpandTableColumn(InvokeEngine, "CleanData", 
        {"Brand", "Channel_Name", "Data_Format_Type", "Forecast_Date", "Value"}
    ),
    
    // 8. Select and order the final columns for your master flat file
    FinalTable = Table.SelectColumns(ExpandedData, {
        "Year Folder", "Month Folder", "Affiliate", "Growth_Driver_Name", "Brand", 
        "File_Type", "Demand_Plan_Month", "Actuals_Month", "Version", 
        "Channel_Name", "Data_Format_Type", "Forecast_Date", "Value"
    })
in
    FinalTable
