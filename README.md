# R365 ODBC Weekly Cash Flow Power Query Work Sample

## Project Title
**Excel Power Query + R365 ODBC Weekly GL Cash Flow Data Pipeline**

## Overview
This work sample demonstrates a refreshable Excel Power Query solution designed for a Restaurant365 / R365 ODBC workflow. The objective is to pull weekly actual GL transaction data from R365, filter it for cash-affecting transactions, map GL accounts into a predefined cash flow reporting structure, and output the results into a clean `R365 Raw Pull` worksheet for later comparison against a Cash Flow Budget.

The workbook is built as a safe demonstration version using mock GL data, but it includes a live ODBC-ready Power Query M script that can be adapted to a real R365 ODBC DSN once credentials and schema details are supplied.

## Business Requirement Simulated
The sample mirrors a client request for an Excel Power Query specialist who can:

- Configure an R365 ODBC connection in an existing workbook
- Pull GL transaction lines from R365
- Filter for a specific entity: `Humble Baron`
- Apply cash-affecting transaction logic
- Pull all weeks from a starting date through a user-selected `As-of Sunday`
- Output the result to a predefined `R365 Raw Pull` tab
- Fully overwrite the raw output on every refresh
- Validate the output against R365 Statement of Cash Flows totals for four historical weeks within a 1% tolerance
- Provide setup, troubleshooting, and user training documentation

## Included Files

| File | Purpose |
|---|---|
| `R365_Power_Query_Work_Sample.xlsx` | Main Excel workbook demo with parameters, mock GL data, mapping, output, and validation tabs |
| `R365_RawPull_Live_ODBC.pq` | Power Query M script template for live R365 ODBC connection |
| `R365_RawPull_Mock_Demo.pq` | Power Query M script using mock workbook data for safe testing |
| `R365_PQ_Setup_Troubleshooting_OnePager.docx` | One-page setup, refresh, validation, and troubleshooting guide |
| `README.md` | Project explanation and usage guide |

## Workbook Tabs

### 1. `Parameters`
Stores user-controlled inputs used by the Power Query process.

Example parameters:

- `StartDate`
- `AsOfSunday`
- `EntityName`
- `ODBC_DSN`

These parameters allow the report user to refresh data for a selected date range and entity without editing the Power Query code.

### 2. `Mock_R365_GL_Lines`
Contains safe mock GL transaction lines used for demonstration. This replaces live R365 data for portfolio/sample purposes.

Typical fields include:

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
Maps GL accounts to the desired cash flow output categories.

Example mapping fields:

- GL Account Number
- GL Account Name
- Cash Flow Section
- Output Column
- Include Flag

### 4. `Cash_Filter_Spec`
Documents the cash-affecting transaction filter logic used by the query.

Example filters include:

- Include only the selected entity
- Include only transactions marked cash-affecting
- Exclude non-cash journal entries
- Exclude accrual-only transactions
- Include dates from `StartDate` through `AsOfSunday`

### 5. `R365 Raw Pull`
Final output tab designed to match the required schema for downstream comparison against the Cash Flow Budget.

The query is designed to fully overwrite this output on each refresh.

### 6. `Validation`
Four-week reconciliation check that compares the Power Query output against expected R365 Statement of Cash Flows totals.

The validation logic flags results as:

- `PASS` when variance is within 1%
- `REVIEW` when variance exceeds the 1% tolerance

## Power Query Logic
The Power Query solution performs the following steps:

1. Reads user parameters from the workbook
2. Connects to R365 through ODBC or mock demo data
3. Pulls GL transaction lines
4. Filters by entity, date range, and cash-affecting rules
5. Joins GL transactions to the GL mapping table
6. Groups the data by week, GL account, and output category
7. Calculates signed cash flow amounts
8. Outputs a clean raw pull table to `R365 Raw Pull`
9. Supports full overwrite behavior during refresh

## Live ODBC Connection Template
The live script uses an ODBC connection pattern similar to:

```powerquery
Odbc.DataSource("dsn=" & ODBC_DSN, [HierarchicalNavigation=true])
```

In a real client environment, the final table names, schema names, and field names would be aligned with the client’s R365 ODBC configuration.

## Refresh Workflow

1. Open the workbook in Excel Desktop.
2. Confirm that the R365 ODBC driver and DSN are configured.
3. Update the `Parameters` tab:
   - Start date
   - As-of Sunday
   - Entity name
   - ODBC DSN
4. Open Power Query Editor.
5. Confirm that the live query points to the correct R365 GL table/view.
6. Refresh the query.
7. Review the `R365 Raw Pull` tab.
8. Review the `Validation` tab for four-week reconciliation status.

## Validation Approach
The validation tab compares weekly Power Query output to expected R365 Statement of Cash Flows totals.

Validation fields include:

- Week Ending Sunday
- PQ Output Total
- R365 Statement of Cash Flows Total
- Variance Amount
- Variance %
- Status

A result is considered successful when the absolute variance percentage is less than or equal to 1%.

## Key Skills Demonstrated

- Excel Power Query development
- Power Query M scripting
- ODBC connection design
- Restaurant365 / R365 data extraction approach
- GL transaction filtering
- Cash-affecting transaction logic
- Parameter-driven reporting
- GL account mapping
- Weekly cash flow data preparation
- Refreshable Excel workbook design
- Data validation and reconciliation
- Client setup and troubleshooting documentation
- User training documentation

## Assumptions
Because this is a work sample, it uses mock data rather than a real client R365 database. The live ODBC script is structured as a production-ready template, but final deployment would require:

- Client R365 ODBC credentials
- Confirmed DSN name
- Confirmed R365 table/view names
- Final GL account mapping file
- Final cash-affecting transaction rules
- Final target schema for `R365 Raw Pull`

## Production Deployment Checklist

- Confirm R365 ODBC driver installation
- Confirm DSN works in Windows ODBC Data Sources
- Confirm user has read access to GL transaction tables/views
- Confirm entity naming matches R365 exactly
- Confirm date fields and posting status fields
- Confirm cash-affecting transaction logic with finance team
- Confirm GL account mapping table
- Load query to `R365 Raw Pull`
- Set refresh to full overwrite
- Reconcile four historical weeks to R365 Statement of Cash Flows
- Train user on parameters, refresh, and validation checks

## Notes for Reviewers
This sample was created to show practical ability to design the data pipe requested in the job description. It is not meant to expose or require real R365 credentials. The workbook demonstrates the structure, logic, documentation, validation workflow, and Power Query approach that would be used in a real client engagement.
