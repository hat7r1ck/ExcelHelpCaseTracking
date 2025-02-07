# Data_Import Sheet Instructions

## Overview

The **Data_Import** sheet serves as the repository for raw case data exported from your case management system (e.g., XSOAR). This sheet is essential as it provides the source data that the VBA macro uses to log help cases in the **HelpCaseLog** sheet. Data can be imported manually or, preferably, automatically using Power Query.

## Expected Data Format

The Data_Import sheet should contain at least the following columns (starting in Column A):

- **Column A: CaseID**  
  - A unique identifier for each case.
  
- **Column B: TimeCreated**  
  - The timestamp when the case was created.
  
- **Column C: OtherInfo**  
  - (Optional) Any additional data or notes related to the case.
  
- **Column D: TimeClosed**  
  - The timestamp when the case was closed.

*Note:* Adjust or add columns if your case data includes more fields. Just be sure to update the VBA code in `Module1.bas` accordingly.

## Setting Up Data Import with Power Query

Using Power Query allows the Data_Import sheet to refresh automatically from an external data source, reducing manual data entry. Follow these steps to set it up:

1. **Prepare Your Data Source:**
   - Export your case data from your case management system as a CSV or Excel file.
   - Save the exported file in a known folder on your computer.

2. **Import Data into Excel:**
   - Open your workbook (`ExcelHelpCaseTracking.xlsm`).
   - Navigate to the **Data** tab on the Excel ribbon.
   - Click on **Get Data** → **From File** → choose **From Text/CSV** (or **From Workbook** if using an Excel file).
   - Browse to the location of your exported file and select it.
   - In the preview dialog, verify that the data is loaded correctly, then click **Load** to import the data into the Data_Import sheet.

3. **Format as a Table:**
   - Once the data is imported, select the entire range.
   - Go to **Insert** → **Table** and create a table.  
     - Give the table a meaningful name (for example, `XSOARData`).
   - Formatting as a table helps ensure that the data range updates automatically when new data is imported.

4. **Configuring Refresh Options (Optional):**
   - With your table selected, go to **Data** → **Queries & Connections**.
   - Right-click your query and choose **Properties**.
   - Set the query to refresh on file open or at regular intervals as needed.

## Customizing the Data_Import Sheet

- **Additional Columns:**  
  If your exported data includes extra fields, add the corresponding columns.  
  Make sure to update any VBA code or formulas that reference column positions.
  
- **Data Validation:**  
  Consider adding data validation or conditional formatting to highlight errors or incomplete data.
  
- **Refreshing Data:**  
  Ensure the Power Query is set up to automatically refresh the data so that the latest case details are always available.

## Troubleshooting

- **Data Not Refreshing:**
  - Verify that the source file path has not changed.
  - Ensure that the query is configured to refresh on file open or manually refresh by clicking **Data → Refresh All**.
  
- **Incorrect Data Format:**
  - Check that the exported file from your case management system matches the expected format (i.e., columns and data types).
  - Adjust the query settings or data transformation steps in Power Query if necessary.
