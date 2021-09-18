// Copyright (c) 2021 Reports Now, Inc.
section mobie;

[DataSource.Kind="mobie", Publish="mobie.Publish"]
shared mobie.Contents = Value.ReplaceType(mobieImpl, mobieType);

mobieImpl = (DatabaseType as text, ServerName as text, Database as text, DataBoxSchemaName as text, FDASPROPName as text) =>
let
		Provider =
			if DatabaseType = "SQL Server" then
				"System.Data.SqlClient"
			else if DatabaseType = "Oracle" then
				"Oracle.DataAccess.Client"
			else
				error "Unknown database type",
		ConnectionString =
			if DatabaseType = "SQL Server" then
				"data source=" & ServerName & ";initial catalog=" & Database
			else
				"data source=" & ServerName,
		FDASPROPResolved =
			if Text.Length(Text.Trim(FDASPROPName)) = 0 then
				"FDASPROP"
			else
				FDASPROPName,
  		dbo_FDASPROP = AdoDotNet.Query(Provider, ConnectionString, "select * from " & DataBoxSchemaName & "." & FDASPROPResolved),
		TableFilterOption = try AdoDotNet.Query(Provider, ConnectionString, "select Value from rn__Options where Name = 'PowerBIShowAllTables'"),
		MobieIdentity = GetMobieIdentity(Provider, ConnectionString, DatabaseType, dbo_FDASPROP),
		DataBoxTableRows = Table.SelectRows(dbo_FDASPROP, each ([DPINTEGER] = 42))[[DPAUSR],[DPFBDINIKY],[DPCDNAME],[DPGTBLOB]],
		DataBoxTableRowsFiltered =
			if TableFilterOption[HasError] = false and Table.RowCount(TableFilterOption[Value]) = 1 and TableFilterOption[Value]{0}[Value] = "1" then
				Table.RemoveColumns(DataBoxTableRows, {"DPAUSR"})
			else
				Table.SelectRows(DataBoxTableRows, each ([DPAUSR] = MobieIdentity))[[DPFBDINIKY],[DPCDNAME],[DPGTBLOB]],
		RenamedDataBoxTableColumns = Table.RenameColumns(DataBoxTableRowsFiltered,{{"DPFBDINIKY","DataBoxSnapshot"},{"DPCDNAME","ReportId"},{"DPGTBLOB","DataBoxTableBlob"}}),
		ExtractedSnapshotName = Table.SplitColumn(RenamedDataBoxTableColumns, "DataBoxSnapshot", Splitter.SplitTextByEachDelimiter({"#(tab)"}, QuoteStyle.Csv, false), {"SnapshotName", "DataBoxName"}),
		TrimmedNames = Table.TransformColumns(ExtractedSnapshotName,{{"DataBoxName", Text.Trim, type text},{"ReportId", Text.Trim, type text}}),
		FilteredDataBoxTableRows = TrimmedNames,
		Decode = (_) => BinaryFormat.Text(Binary.Length(_))(_),
		ReportDataSourceRows = Table.SelectRows(dbo_FDASPROP, each ([DPINTEGER] = 49))[[DPFBDINIKY],[DPCDNAME],[DPGTBLOB]],
		RenamedReportDataSourceColumns = Table.RenameColumns(ReportDataSourceRows,{{"DPFBDINIKY", "ReportName"}, {"DPCDNAME", "ReportId"}, {"DPGTBLOB", "ReportDataSourceBlob"}}),
		TrimmedReportNames = Table.TransformColumns(RenamedReportDataSourceColumns,{{"ReportName", Text.Trim, type text},{"ReportId", Text.Trim, type text}}),
		DecodedReportDataSourceBlob = Table.TransformColumns(TrimmedReportNames, {{"ReportDataSourceBlob", each Xml.Tables(Decode(Binary.Decompress(_, Compression.GZip)))}}),
		ExpandedReportDataSourceTopProperty = Table.ExpandTableColumn(DecodedReportDataSourceBlob, "ReportDataSourceBlob", {"property"}, {"ReportDataSourceBlob.property"}),
		ExpandedReportDataSourceXmlProperties = Table.ExpandTableColumn(ExpandedReportDataSourceTopProperty, "ReportDataSourceBlob.property",{"Element:Text", "Attribute:name"}, {"ReportDataSourceAttributeValue", "ReportDataSourceAttributeName"}),
		RemovedReportDataSourceColumns = Table.RemoveColumns(ExpandedReportDataSourceXmlProperties, {"ReportDataSourceAttributeName"}),
		RenamedReportDataSourceProperties = Table.RenameColumns(RemovedReportDataSourceColumns,{{"ReportDataSourceAttributeValue","DataName"}}),
		JoinedRows = Table.Join(FilteredDataBoxTableRows, "ReportId", RenamedReportDataSourceProperties, "ReportId"),
		DistinctJoinedRows = Table.Distinct(JoinedRows, { "ReportName", "DataName" }),
		DecodedDataBoxTableBlob = Table.TransformColumns(DistinctJoinedRows, {{"DataBoxTableBlob", each Xml.Tables(Decode(Binary.Decompress(_, Compression.GZip)))}}),
		ExpandedDataBoxTableTopProperty = Table.ExpandTableColumn(DecodedDataBoxTableBlob, "DataBoxTableBlob", {"property"}, {"TopProperty"}),
		ExpandedDataBoxTableXmlProperties = Table.ExpandTableColumn(ExpandedDataBoxTableTopProperty, "TopProperty", {"Element:Text", "Attribute:name", "property", "Attribute:value"}, {"PropertyText", "PropertyAttributeName", "PropertyNestedValue", "PropertyAttributeValue"}),
		TableIds = Table.SelectRows(ExpandedDataBoxTableXmlProperties, each [PropertyAttributeName] = "TableId"),
		RenamedDataBoxTableProperties = Table.RenameColumns(TableIds, {{"PropertyText","TableId"}}),
		RnTableIds = Table.RemoveColumns(RenamedDataBoxTableProperties,{"PropertyAttributeName", "PropertyNestedValue", "PropertyAttributeValue"}),
		NavTable = CreateNavTable(RnTableIds, ExpandedDataBoxTableXmlProperties, Provider, ConnectionString, DataBoxSchemaName, FDASPROPResolved)
in
    NavTable;

mobieType = type function (
	DatabaseType as (type text meta [
		Documentation.FieldCaption = "Database Type",
		Documentation.AllowedValues = { "SQL Server", "Oracle" }
	]),
	ServerName as (type text meta [
		Documentation.FieldCaption = "Server Name",
		Documentation.FieldDescription = { "DataBox database server name for connection string" }
	]),
	Database as (type text meta [
		Documentation.FieldCaption = "Database Name",
		Documentation.SampleValues = { "DataBox" }
	]),
	DataBoxSchemaName as (type text meta [
		Documentation.FieldCaption = "Schema Name",
		Documentation.SampleValues = { "dbo" }
	]))
	as table meta [
		Documentation.Name = "mobie® DataBox",
		Documentation.LongDescription = "Returns a table of all DataBox tables stored in the specified DataBox database"
	];
		
GetMobieIdentity = (Provider as text, ConnectionString as text, DatabaseType as text, dbo_FDASPROP as table) =>
    let
  		SessionUser =
			if DatabaseType = "SQL Server" then
				AdoDotNet.Query(Provider, ConnectionString, "select UPPER(TRIM(SYSTEM_USER)) as UserName"){0}[UserName]
			else
				AdoDotNet.Query(Provider, ConnectionString, "select UPPER(TRIM(SYS_CONTEXT('USERENV', 'SESSION_USER'))) as USERNAME FROM DUAL"){0}[USERNAME],
		MobieIdentityRows = Table.SelectRows(dbo_FDASPROP, each ([DPINTEGER] = 53 and Text.Trim([DPFBDINIKY]) = SessionUser))[[DPAUSR]],
		MobieIdentity =
			if Table.RowCount(MobieIdentityRows) = 0 then 
				error "No matching mobie identity found for " & SessionUser
			else
				MobieIdentityRows{0}[DPAUSR]
    in
        MobieIdentity;

GetSecurityAssociations = (Provider as text, ConnectionString as text, DataBoxSchemaName as text) as table =>
    let
        dbo_FDASPROP = AdoDotNet.Query(Provider, ConnectionString, "select * from " & DataBoxSchemaName & ".FDASPROP"),
        MobieUserRecords = Table.SelectRows(dbo_FDASPROP, each ([DPINTEGER] = 52 and [DPAUSR] <> "DAS_PUBLIC"))[[DPAUSR],[DPGTBLOB]],
        Decode = (_) => BinaryFormat.Text(Binary.Length(_))(_),
        DecodedMobieUserBlob = Table.TransformColumns(MobieUserRecords, {{"DPGTBLOB", each Xml.Tables(Decode(Binary.Decompress(_, Compression.GZip)))}}),
        UserAndSecurityAssociationTable = Table.TransformColumns(DecodedMobieUserBlob, {{"DPGTBLOB", each Table.SelectRows(_[property]{0}, each ([#"Attribute:name"] = "SecurityAssociations"))[property]{0}}}),
        ExpandedAssociations = Table.ExpandTableColumn(UserAndSecurityAssociationTable,"DPGTBLOB",{"property"}),
        DbNullsHandled = Table.TransformColumns(ExpandedAssociations, {{"property", each Table.CombineColumns(_,{"Element:Text","Attribute:type"}, (ColList) => if ColList{1} = "System.DBNull" then null else ColList{0}, "Value")}}),
        PivotedAssociations = Table.TransformColumns(DbNullsHandled, {{"property", each Table.Pivot(_,{"Name","TypeName","Value"},"Attribute:name","Value")}}),
        UserAndSecurityAssociations = Table.ExpandTableColumn(PivotedAssociations,"property",{"Name","TypeName","Value"}),
        MobieIdentityRecords = Table.SelectRows(dbo_FDASPROP, each ([DPINTEGER] = 53))[[DPAUSR],[DPFBDINIKY]],
        IdentityAndSecurityAssociations = Table.Join(UserAndSecurityAssociations, "DPAUSR", MobieIdentityRecords, "DPAUSR"),
        SecurityAssociationsRenamed = Table.RenameColumns(IdentityAndSecurityAssociations,{{"DPFBDINIKY", "Identity"},{"Name","SecurityAssociation"}}),
        SecurityAssociationsUnpivoted = Table.RemoveColumns(SecurityAssociationsRenamed, {"DPAUSR", "TypeName"}),
        SecurityAssociations = Table.Pivot(SecurityAssociationsUnpivoted, List.Distinct(Table.Column(SecurityAssociationsUnpivoted, "SecurityAssociation")), "SecurityAssociation", "Value")
    in
        SecurityAssociations;

GetDataBoxTable = (reportId as text, tableId as text, dataBoxProperties as table, Provider as text, ConnectionString as text, DataBoxSchemaName as text) as table =>
	let
		base = AdoDotNet.Query(Provider, ConnectionString, "select * from " & DataBoxSchemaName & ".rn_" & tableId),
        FilteredDataBoxProperties = Table.SelectRows(dataBoxProperties, each [ReportId] = reportId),
		FilteredColumnTranslations = Table.SelectRows(FilteredDataBoxProperties, each [PropertyAttributeName] = "ColumnTranslations"),
		ColumnTranslationsPlus = Table.ExpandTableColumn(FilteredColumnTranslations, "PropertyNestedValue", {"property"},{"ColumnTranslation"}),
		ColumnTranslations = Table.RemoveColumns(ColumnTranslationsPlus, {"SnapshotName", "DataBoxName", "ReportId", "PropertyText", "PropertyAttributeName", "PropertyAttributeValue", "ReportName", "DataName" }),
		SymbolMapperPivot = (_) => Table.ToRecords(Table.Pivot(_, {"FromName", "ToName"}, "Attribute:name", "Element:Text")){0},
		PivotedColumnTranslations = Table.TransformColumns(ColumnTranslations, {{ "ColumnTranslation", SymbolMapperPivot}}),
		ExpandedColumnTranslations = Table.ExpandRecordColumn(PivotedColumnTranslations, "ColumnTranslation", {"FromName", "ToName"}),
		ColumnTranslationsRenamed = Table.RenameColumns(ExpandedColumnTranslations, {{"ToName", "ColumnName"}}),

		FilteredCaptionTranslations = Table.SelectRows(FilteredDataBoxProperties, each [PropertyAttributeName] = "FieldNameCaptionTranslations"),
		CaptionTranslationsPlus = Table.ExpandTableColumn(FilteredCaptionTranslations, "PropertyNestedValue", {"property"},{"FieldNameCaptionTranslation"}),
		CaptionTranslations = Table.RemoveColumns(CaptionTranslationsPlus, {"SnapshotName", "DataBoxName", "ReportId", "PropertyText", "PropertyAttributeName", "PropertyAttributeValue", "ReportName", "DataName" }),
		PivotedCaptionTranslations = Table.TransformColumns(CaptionTranslations, {{ "FieldNameCaptionTranslation", SymbolMapperPivot}}),
		ExpandedCaptionTranslations = Table.ExpandRecordColumn(PivotedCaptionTranslations, "FieldNameCaptionTranslation", {"FromName", "ToName"}),
		CaptionTranslationsRenamed = Table.RenameColumns(ExpandedCaptionTranslations, {{"ToName", "Caption"}}),

		ColumnCaptionMapping = Table.Join(ColumnTranslationsRenamed, "FromName", CaptionTranslationsRenamed, "FromName"),
		SymbolMapper = (column as text) => let MappedName = Table.SelectRows(ColumnCaptionMapping, each Comparer.OrdinalIgnoreCase([ColumnName], column) = 0) in if Table.RowCount(MappedName) = 0 then column else MappedName{0}[Caption],

		FinalTable = Table.TransformColumnNames(base, SymbolMapper)  
	in
	    FinalTable;

CreateNavTable = (base as table, dataBoxProperties as table, Provider as text, ConnectionString as text, DataBoxSchemaName as text, FDASPROPName as text) as table =>
(base as table, dataBoxProperties as table, Provider as text, ConnectionString as text, DataBoxSchemaName as text) as table =>
    let
		withData = Table.AddColumn(base, "Data", each (GetDataBoxTable([ReportId], [TableId], dataBoxProperties, Provider, ConnectionString, DataBoxSchemaName)), type table),
		withNames = Table.CombineColumns(withData, { "ReportName", "DataName", "SnapshotName" }, Combiner.CombineTextByDelimiter(", "), "Name"),
		change1 = Table.RemoveColumns(withNames, {"DataBoxName", "TableId", "ReportId"}),
        wrappedAssociations = Table.FromRecords({[Name = "Security Associations", Data = GetSecurityAssociations(Provider, ConnectionString, DataBoxSchemaName, FDASPROPName)]}),
        withAssociations = Table.InsertRows(change1, 0, wrappedAssociations),
        withItemKind = Table.AddColumn(withAssociations, "ItemKind", each "Table", type text),
        withItemName = Table.AddColumn(withItemKind, "ItemName", each "Table", type text),
        withIsLeaf = Table.AddColumn(withItemName, "IsLeaf", each true, type logical),
 		navTable = Table.ToNavigationTable(withIsLeaf, {"Name"}, "Name", "Data", "ItemKind", "ItemName", "IsLeaf")
    in
        navTable;
		
Table.ToNavigationTable = (
    table as table,
    keyColumns as list,
    nameColumn as text,
    dataColumn as text,
    itemKindColumn as text,
    itemNameColumn as text,
    isLeafColumn as text
) as table =>
    let
        tableType = Value.Type(table),
        newTableType = Type.AddTableKey(tableType, keyColumns, true) meta 
        [
            NavigationTable.NameColumn = nameColumn, 
            NavigationTable.DataColumn = dataColumn,
            NavigationTable.ItemKindColumn = itemKindColumn, 
            Preview.DelayColumn = itemNameColumn, 
            NavigationTable.IsLeafColumn = isLeafColumn
        ],
        navigationTable = Value.ReplaceType(table, newTableType)
    in
        navigationTable;

// Data Source Kind description
mobie = [
	TestConnection = (dataSourcePath) =>
		let
			json = Json.Document(dataSourcePath),
			DatabaseType = json[DatabaseType],
			ServerName = json[ServerName],
			Database = json[Database],
			DataBoxSchemaName = json[DataBoxSchemaName],
			FDASPROPName = json[FDASPROPName]
		in
			{ "mobie.Contents", DatabaseType, ServerName, Database, DataBoxSchemaName, FDASPROPName },
    Authentication = [
        UsernamePassword = [],
        Windows = []
    ],
	// See https://docs.microsoft.com/en-us/power-query/handlingresourcepath
    Label = Extension.LoadString("DataSourceLabel")
];

// Data Source UI publishing description
mobie.Publish = [
    Beta = false,
    Category = "Database",
    ButtonText = { Extension.LoadString("ButtonTitle"), Extension.LoadString("ButtonHelp") },
    LearnMoreUrl = "https://reportsnow.com/products/mobie",
    SupportsDirectQuery = false,
    SourceImage = mobie.Icons,
    SourceTypeImage = mobie.Icons
];

mobie.Icons = [
    Icon16 = { Extension.Contents("mobie16.png"), Extension.Contents("mobie20.png"), Extension.Contents("mobie24.png"), Extension.Contents("mobie32.png") },
    Icon32 = { Extension.Contents("mobie32.png"), Extension.Contents("mobie40.png"), Extension.Contents("mobie48.png"), Extension.Contents("mobie64.png") }
];
