# R365 ODBC Weekly Cash Flow Power Query Work

## Project Title
**Excel Power Query + R365 ODBC Weekly GL Cash Flow Data Pipeline**

## Start Here: Main Output File
The primary deliverable in this work sample is:

**[`R365_Power_Query_Work_Sample.xlsx`](./R365_Power_Query_Work_Sample.xlsx)**

This Excel workbook is the main output file. It demonstrates the full data-pipe structure requested in the Upwork job description, including user parameters, mock R365 GL data, GL account mapping, cash-affecting filter logic, the final `R365 Raw Pull` output tab, and a four-week validation/reconciliation tab.

## Work Sample Summary
This sample demonstrates how I would build a refreshable Excel Power Query solution that connects to Restaurant365 / R365 through ODBC, extracts weekly GL transaction lines, applies entity/date/cash-affecting filters, maps GL accounts into a cash flow reporting structure, and outputs a clean raw data table for later comparison against a Cash Flow Budget.

The workbook uses safe mock data for portfolio review, but it also includes a live ODBC-ready Power Query M script that can be adapted to the client’s real R365 DSN, credentials, schema, and table names.

## Job Requirement Covered
This work sample mirrors the requested scope:

- Configure an R365 ODBC connection in an Excel workbook
- Pull GL transaction lines from R365
- Filter for the entity `Humble Baron`
- Apply cash-affecting transaction logic
- Pull all weeks from a starting date through a user-selected `As-of Sunday`
- Output the data to a predefined `R365 Raw Pull` tab
- Fully overwrite the output on each refresh
- Validate four historical weeks against R365 Statement of Cash Flows totals within a 1% threshold
- Provide setup, troubleshooting, and user training documentation

## Included Files

| File | Purpose |
|---|---|
| [`R365_Power_Query_Work_Sample.xlsx`](./R365_Power_Query_Work_Sample.xlsx) | Main Excel workbook output file with parameters, mock data, mapping, filter logic, final raw pull output, and validation |
| [`R365_RawPull_Live_ODBC.pq`](./R365_RawPull_Live_ODBC.pq) | Production-style Power Query M template for connecting to live R365 through ODBC |
| [`R365_RawPull_Mock_Demo.pq`](./R365_RawPull_Mock_Demo.pq) | Mock/demo Power Query M script that can be reviewed without live R365 credentials |
| [`R365_PQ_Setup_Troubleshooting_OnePager.docx`](./R365_PQ_Setup_Troubleshooting_OnePager.docx) | One-page setup, refresh, troubleshooting, and user training document |
| [`README.md`](./README.md) | Project overview and reviewer guide |

## Excel Workbook Structure

### 1. `Parameters`
Stores user-controlled values used by the query.

Typical fields:

- `StartDate`
- `AsOfSunday`
- `EntityName`
- `ODBC_DSN`

These allow the user to change the reporting window and entity without editing the M code.

### 2. `Mock_R365_GL_Lines`
Contains sample GL transaction lines used for safe review. In production, this table would be replaced by the live R365 ODBC source.

Example fields:

- Entity
- Transaction Date
- Week Ending Sunday
- GL Account Number
- GL Account Name
- Debit
- Credit
- Net Amount
- Transaction Type
- Cash Affecting Flag
- Description
- Vendor/Customer
- Source Module

### 3. `GL_Mapping`
Maps GL accounts to the cash flow output structure.

Example fields:

- GL Account Number
- GL Account Name
- Cash Flow Section
- Output Column
- Include Flag

### 4. `Cash_Filter_Spec`
Documents the cash-affecting transaction rules used by the query.

Example rules:

- Include only the selected entity
- Include only transactions marked as cash-affecting
- Exclude non-cash journal entries
- Exclude accrual-only transactions
- Include transactions from `StartDate` through `AsOfSunday`

### 5. `R365 Raw Pull`
This is the final output tab requested in the job description.

The query output is designed to fully overwrite this tab on every refresh. It contains the cleaned weekly ACTUAL cash flow GL data that can later be compared with the Cash Flow Budget.

### 6. `Validation`
Compares four historical weeks of Power Query output against expected R365 Statement of Cash Flows totals.

Validation fields include:

- Week Ending Sunday
- Power Query Output Total
- R365 Statement of Cash Flows Total
- Variance Amount
- Variance %
- Status

A week is marked `PASS` when the absolute variance percentage is less than or equal to 1%. Anything above 1% is marked for review.

## Power Query Process
The Power Query workflow follows this sequence:

1. Read parameter values from the workbook
2. Connect to R365 through ODBC, or use mock workbook data for demo mode
3. Pull GL transaction lines
4. Filter by entity
5. Filter by transaction date range
6. Apply cash-affecting transaction logic
7. Join GL transaction lines to the GL mapping table
8. Remove excluded GL accounts or transaction categories
9. Group and organize weekly data
10. Output the final clean result to `R365 Raw Pull`
11. Refresh as a full overwrite table
12. Validate four historical weeks against expected R365 Statement of Cash Flows totals

## Live ODBC Connection Pattern
The included live M script uses an ODBC connection pattern similar to:

```powerquery
Odbc.DataSource("dsn=" & ODBC_DSN, [HierarchicalNavigation=true])
```

In a real client workbook, the final query would be aligned to the client’s actual R365 ODBC schema, including the correct database, schema, table/view names, field names, posting status fields, and transaction type fields.

## Setup and Refresh Workflow

1. Open `R365_Power_Query_Work_Sample.xlsx` in Excel Desktop.
2. Confirm the R365 ODBC driver is installed.
3. Confirm the R365 DSN is configured in Windows ODBC Data Sources.
4. Update the `Parameters` tab.
5. Open Power Query Editor.
6. Replace the mock source with the live ODBC source where needed.
7. Confirm the GL transaction table/view and field names.
8. Refresh the query.
9. Review the `R365 Raw Pull` tab.
10. Review the `Validation` tab and confirm the four-week reconciliation status.

## Production Deployment Checklist

- Confirm R365 ODBC driver installation
- Confirm DSN name and authentication method
- Confirm user read access to GL transaction tables/views
- Confirm exact entity name for `Humble Baron`
- Confirm start date and selected `As-of Sunday`
- Confirm transaction date/posting date logic
- Confirm cash-affecting transaction filter specification
- Confirm GL Account → output column mapping table
- Confirm target schema for `R365 Raw Pull`
- Load final query output to the scaffolded tab
- Set refresh behavior to full overwrite
- Reconcile four historical weeks within 1%
- Train the user on parameters, refresh, validation, and troubleshooting

## Skills Demonstrated

- Excel Power Query development
- Power Query M scripting
- R365 / Restaurant365 ODBC data extraction approach
- Parameter-driven workbook design
- GL transaction filtering
- Cash-affecting transaction logic
- GL account mapping
- Weekly cash flow raw data preparation
- Full-refresh output table design
- Four-week validation and reconciliation
- Setup and troubleshooting documentation
- End-user training handoff

## Assumptions
This work sample does not include real R365 credentials or live client data. It uses mock GL transaction data so the logic can be reviewed safely. For production use, the live ODBC script would be finalized after receiving:

- R365 ODBC credentials
- DSN name
- R365 table/view names
- GL mapping table
- Cash-affecting filter specification
- Target output schema
- Sample expected output
- Four historical weeks for reconciliation

## Reviewer Note
This package is intended to show that I can handle the requested R365 Power Query data-pipe work: connection design, parameter handling, filtering, mapping, output formatting, overwrite refresh behavior, validation, troubleshooting documentation, and user training.
