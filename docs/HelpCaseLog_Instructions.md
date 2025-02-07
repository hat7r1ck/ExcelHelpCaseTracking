# HelpCaseLog Sheet Instructions

## Overview

The **HelpCaseLog** sheet is the central log that stores every help case entry recorded by the QuickEntry VBA macro. This sheet provides a historical record of cases where assistance was provided, capturing key details such as the case ID, creation time, help timestamp, closure time, and any optional notes. This data can be used for reporting, tracking performance metrics, and verifying that all help events are logged.

## Sheet Layout

The HelpCaseLog sheet is structured with the following columns:

- **Column A: CaseID**  
  - The unique identifier for the case. This is pulled from the Data_Import sheet.
  
- **Column B: TimeCreated**  
  - The original creation time of the case, as provided by the Data_Import sheet.
  
- **Column C: HelpTimestamp**  
  - The timestamp when the help case was logged (the current time when the VBA macro runs).
  
- **Column D: TimeClosed**  
  - The closure time of the case, as provided by the Data_Import sheet.
  
- **Column E: Notes**  
  - Optional notes or comments entered via the QuickEntry sheet.

## How It Works

1. **Logging a Help Case:**  
   When a help case is logged via the QuickEntry sheet (by entering a CaseID in cell B2 and clicking the "Add Help Case" button), the VBA macro:
   - Searches for the corresponding CaseID in the Data_Import sheet.
   - Retrieves related details such as the TimeCreated and TimeClosed.
   - Appends a new row in the HelpCaseLog sheet with these details.
   - Records the current time as the HelpTimestamp.
   - Adds any optional notes provided.

2. **Automatic Updates:**  
   Each time the macro runs, a new row is added to HelpCaseLog, ensuring that every help event is captured for future review or reporting.

## Customizing the HelpCaseLog Sheet

- **Formatting as a Table:**  
  It is recommended to format your HelpCaseLog data as an Excel Table. This allows for easier sorting, filtering, and dynamic referencing in PivotTables or dashboards.
  
- **Adding Extra Columns:**  
  If additional details need to be captured (for example, a status or category field), you can add new columns to the sheet. Make sure to update the VBA code in `Module1.bas` to populate these extra columns.
  
- **Styling and Conditional Formatting:**  
  Customize the appearance of the HelpCaseLog sheet by applying your preferred styles and conditional formatting rules. This can help highlight important data or anomalies in the log.

## Troubleshooting

- **Missing Entries:**  
  If you notice that a help case isn't logged:
  - Verify that the correct CaseID was entered in the QuickEntry sheet.
  - Check that the Data_Import sheet contains the corresponding case details.
  - Ensure that macros are enabled and the VBA code executed without errors.

- **Data Misalignment:**  
  If data appears in the wrong columns:
  - Confirm that the cell references in the VBA code match the layout of the HelpCaseLog sheet.
  - Adjust the column mapping in the code if you have modified the sheet structure.

- **Incorrect Data Logging:**  
  Ensure that the data in the Data_Import sheet is accurate. Any discrepancies in the source data will be reflected in the log.
