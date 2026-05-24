// Query name: R365 Raw Pull - Mock Demo
// Purpose: Refreshes from the included Mock_GL_Lines table, so the workbook can be reviewed without R365 access.

let
    ParamTable = Excel.CurrentWorkbook(){[Name="tblParameters"]}[Content],
    GetParam = (paramName as text) as any =>
        ParamTable{[Parameter = paramName]}[Value],

    pStartDate = Date.From(GetParam("pStartDate")),
    pAsOfSunday = Date.From(GetParam("pAsOfSunday")),
    pEntity = Text.From(GetParam("pEntity")),

    Source = Excel.CurrentWorkbook(){[Name="Mock_GL_Lines"]}[Content],

    Typed = Table.TransformColumnTypes(
        Source,
        {
            {"TransactionDate", type date}, {"WeekEndingSunday", type date}, {"Entity", type text},
            {"GLAccount", type text}, {"AccountName", type text}, {"Department", type text},
            {"VendorPayee", type text}, {"Memo", type text}, {"SourceModule", type text},
            {"JournalEntryNo", type text}, {"Debit", Currency.Type}, {"Credit", Currency.Type},
            {"IsCashAccount", type logical}, {"CashAffectingFlag", type logical}
        }
    ),

    CashOnly = Table.SelectRows(
        Typed,
        each [Entity] = pEntity
            and [TransactionDate] >= pStartDate
            and [TransactionDate] <= pAsOfSunday
            and [IsCashAccount] = true
            and [CashAffectingFlag] = true
            and not Text.Contains(Text.Upper(Text.From([SourceModule])), "TRANSFER")
    ),

    Mapping = Excel.CurrentWorkbook(){[Name="GL_Mapping"]}[Content],
    MappingTyped = Table.TransformColumnTypes(Mapping, {{"GLAccount", type text}, {"CashFlowSection", type text}, {"OutputColumn", type text}}),

    Joined = Table.NestedJoin(CashOnly, {"GLAccount"}, MappingTyped, {"GLAccount"}, "Map", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "Map", {"CashFlowSection", "OutputColumn"}, {"CashFlowSection", "OutputColumn"}),

    NetCash = Table.AddColumn(Expanded, "NetCashAmount", each [Debit] - [Credit], Currency.Type),
    IncludeFlag = Table.AddColumn(NetCash, "IncludeInPull", each true, type logical),
    DataSource = Table.AddColumn(IncludeFlag, "DataSource", each "Mock_GL_Lines / ODBC pattern", type text),

    Final = Table.SelectColumns(
        DataSource,
        {
            "Entity", "WeekEndingSunday", "TransactionDate", "GLAccount", "AccountName",
            "CashFlowSection", "OutputColumn", "VendorPayee", "Memo", "SourceModule",
            "JournalEntryNo", "Debit", "Credit", "NetCashAmount", "IncludeInPull", "DataSource"
        }
    ),

    Sorted = Table.Sort(Final, {{"WeekEndingSunday", Order.Ascending}, {"TransactionDate", Order.Ascending}, {"JournalEntryNo", Order.Ascending}})
in
    Sorted
