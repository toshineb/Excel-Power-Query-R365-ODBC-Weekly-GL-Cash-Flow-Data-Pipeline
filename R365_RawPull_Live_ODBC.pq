// Query name: R365 Raw Pull - Live ODBC
// Purpose: Pull weekly ACTUAL cash-flow GL lines from R365 via ODBC and load to "R365 Raw Pull".
// Replace dbo.GLTransactionLines and field names with the exact R365 ODBC schema supplied at kickoff.

let
    // Excel named-table parameters
    ParamTable = Excel.CurrentWorkbook(){[Name="tblParameters"]}[Content],
    GetParam = (paramName as text) as any =>
        ParamTable{[Parameter = paramName]}[Value],

    pStartDate = Date.From(GetParam("pStartDate")),
    pAsOfSunday = Date.From(GetParam("pAsOfSunday")),
    pEntity = Text.From(GetParam("pEntity")),
    pODBC_DSN = Text.From(GetParam("pODBC_DSN")),

    // Safety check: as-of date should be Sunday
    CheckSunday =
        if Date.DayOfWeek(pAsOfSunday, Day.Sunday) <> 0
        then error "Parameter pAsOfSunday must be a Sunday."
        else pAsOfSunday,

    ODBCSource = Odbc.DataSource("dsn=" & pODBC_DSN, [HierarchicalNavigation = true]),

    Source = Value.NativeQuery(
        ODBCSource,
        "
        SELECT
            CAST(TransactionDate AS date) AS TransactionDate,
            EntityName AS Entity,
            AccountNumber AS GLAccount,
            AccountName,
            DepartmentName AS Department,
            VendorName AS VendorPayee,
            Memo,
            SourceModule,
            JournalEntryNumber AS JournalEntryNo,
            DebitAmount AS Debit,
            CreditAmount AS Credit,
            IsCashAccount,
            CashAffectingFlag
        FROM dbo.GLTransactionLines
        WHERE CAST(TransactionDate AS date) >= ?
          AND CAST(TransactionDate AS date) <= ?
          AND EntityName = ?
        ",
        {pStartDate, CheckSunday, pEntity}
    ),

    Typed = Table.TransformColumnTypes(
        Source,
        {
            {"TransactionDate", type date}, {"Entity", type text}, {"GLAccount", type text},
            {"AccountName", type text}, {"Department", type text}, {"VendorPayee", type text},
            {"Memo", type text}, {"SourceModule", type text}, {"JournalEntryNo", type text},
            {"Debit", Currency.Type}, {"Credit", Currency.Type},
            {"IsCashAccount", type logical}, {"CashAffectingFlag", type logical}
        }
    ),

    CashOnly = Table.SelectRows(
        Typed,
        each [Entity] = pEntity
            and [TransactionDate] >= pStartDate
            and [TransactionDate] <= CheckSunday
            and [IsCashAccount] = true
            and [CashAffectingFlag] = true
            and not Text.Contains(Text.Upper(Text.From([SourceModule])), "TRANSFER")
    ),

    WeekEnding = Table.AddColumn(
        CashOnly,
        "WeekEndingSunday",
        each Date.EndOfWeek([TransactionDate], Day.Sunday),
        type date
    ),

    Mapping = Excel.CurrentWorkbook(){[Name="GL_Mapping"]}[Content],
    MappingTyped = Table.TransformColumnTypes(Mapping, {{"GLAccount", type text}, {"CashFlowSection", type text}, {"OutputColumn", type text}}),

    Joined = Table.NestedJoin(WeekEnding, {"GLAccount"}, MappingTyped, {"GLAccount"}, "Map", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "Map", {"CashFlowSection", "OutputColumn"}, {"CashFlowSection", "OutputColumn"}),

    NetCash = Table.AddColumn(Expanded, "NetCashAmount", each [Debit] - [Credit], Currency.Type),
    IncludeFlag = Table.AddColumn(NetCash, "IncludeInPull", each true, type logical),
    DataSource = Table.AddColumn(IncludeFlag, "DataSource", each "R365 ODBC", type text),

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
