# QuickEntry Sheet Instructions

## Overview

The **QuickEntry** sheet is designed to provide a simple and efficient interface for logging help cases. When an analyst reaches out for assistance, you can quickly enter the case details into this sheet. The sheet then uses a VBA macro (provided in `Module1.bas`) to look up the case in the data import sheet and log the help event (with a timestamp) in the HelpCaseLog sheet.

## Layout of the QuickEntry Sheet

The QuickEntry sheet is structured as follows:

- **Cell B2:**  
  - **Purpose:** Input field for entering the CaseID.
  - **Instructions:** Type the CaseID corresponding to the case for which you provided help.
  
- **Cell B3 (Optional):**  
  - **Purpose:** Input field for entering any additional notes or comments about the case.
  - **Instructions:** If needed, provide any extra information about the help event.
  
- **Add Help Case Button:**  
  - **Purpose:** A clickable button that triggers the VBA macro (`AddHelpCase`).
  - **Instructions:** Once you have entered the required data, click this button to log the help case.

## How to Use the QuickEntry Sheet

Follow these steps to log a help case:

1. **Enter the CaseID:**
   - Navigate to cell **B2** on the QuickEntry sheet.
   - Input the CaseID for the case you assisted with.

2. **Enter Optional Notes:**
   - (Optional) Navigate to cell **B3**.
   - Input any additional notes you want to record along with the help case.

3. **Log the Help Case:**
   - Click the **"Add Help Case"** button.
   - The button is linked to the `AddHelpCase` VBA macro, which will:
     - Retrieve the CaseID from cell B2.
     - Optionally, retrieve any notes from cell B3.
     - Search for the case in the **Data_Import** sheet.
     - Log the help event in the **HelpCaseLog** sheet with the current timestamp.

4. **Confirmation:**
   - A message box will appear confirming that the case has been logged successfully.
   - Verify the entry by checking the **HelpCaseLog** sheet.

## Customizing the QuickEntry Sheet

If you need to make changes to the QuickEntry sheet, consider the following:

- **Cell References:**
  - Ensure that the cell references in the VBA code (e.g., CaseID in B2, Notes in B3) match the layout of your QuickEntry sheet.
  
- **Button Assignment:**
  - To verify or change the macro assignment, right-click on the "Add Help Case" button, select **Assign Macro**, and ensure it is set to `AddHelpCase`.

- **Adding Additional Fields:**
  - If you require extra fields, update both the QuickEntry sheet layout and the corresponding VBA code in `Module1.bas`.

## Troubleshooting

- **Macro Not Running:**
  - Confirm that macros are enabled in Excel. (See [VBA Integration Instructions](./VBA_Integration_Instructions.md) for details on enabling macros.)
  - Verify that the VBA module (Module1) is correctly imported into your workbook.

- **CaseID Not Found:**
  - Double-check that the entered CaseID exists in the **Data_Import** sheet.
  - Ensure that there are no extra spaces or typos in the CaseID.

- **Incorrect Data Logging:**
  - Make sure the cell references in the VBA code align with those on the QuickEntry sheet.
  - Review the HelpCaseLog sheet to verify that the logged data is accurate.
