# VBA Integration Instructions

This document explains how to add the provided VBA code (stored in `Module1.bas`) into your Excel workbook. Follow the steps below to import and integrate the module into your project.

---

## Prerequisites

- **Excel Help Case Tracking Workbook:**  
  Ensure you have your workbook (e.g., `ExcelHelpCaseTracking.xlsm`) saved as a Macro-Enabled Workbook.
  
- **VBA Module File:**  
  The VBA code is stored in the `Module1.bas` file, located in the `src/` folder of the repository.

- **Macros Enabled:**  
  Make sure that macros are enabled in Excel (see "Enabling Macros" below).

---

## Steps to Import the VBA Module

### 1. Open Your Workbook

- Open `ExcelHelpCaseTracking.xlsm` in Microsoft Excel.
- If prompted, enable macros to allow the VBA code to run.

### 2. Open the Visual Basic for Applications (VBA) Editor

- Press <kbd>Alt</kbd> + <kbd>F11</kbd> to open the VBA editor.
- Alternatively, go to the **Developer** tab in the Excel ribbon and click on **Visual Basic**.

### 3. Import the Module

- In the VBA editor, locate your project in the **Project Explorer** (usually on the left side). Your project will be listed as something like `VBAProject (ExcelHelpCaseTracking.xlsm)`.
- Right-click on your project name and select **Import File...**.
- Browse to the location of the `Module1.bas` file (located in the `src/` folder of this repository).
- Select `Module1.bas` and click **Open**.  
  This action adds the module to your project under the **Modules** folder.

### 4. Verify the Imported Code

- In the **Project Explorer**, under your project, expand the **Modules** folder.
- Double-click on **Module1** to open it in the editor.
- Review the code to ensure it matches the content provided in `Module1.bas`.

### 5. Save Your Workbook

- Save your workbook by pressing <kbd>Ctrl</kbd> + <kbd>S</kbd>.
- Make sure it remains saved as a Macro-Enabled Workbook (`.xlsm`).

---

## Enabling Macros

If macros are disabled in your Excel settings, follow these steps to enable them:

1. Go to **File** → **Options**.
2. Select **Trust Center** from the menu on the left.
3. Click on **Trust Center Settings...**.
4. Select **Macro Settings**.
5. Choose **Enable all macros** (or the appropriate setting as per your organization’s policy).
6. Click **OK** and restart Excel if necessary.

---

## Testing the Macro

1. **Navigate to the QuickEntry Sheet:**  
   Go to the **QuickEntry** sheet in your workbook.

2. **Enter Test Data:**  
   Enter a valid CaseID in the designated cell (e.g., cell **B2**).  
   (Optionally, enter any additional notes in cell **B3** if your setup supports it.)

3. **Run the Macro:**  
   - Click on the **Add Help Case** button (which is linked to the `AddHelpCase` macro).
   - Alternatively, you can run the macro directly from the VBA editor by selecting the `AddHelpCase` subroutine and pressing <kbd>F5</kbd>.

4. **Verify the Log:**  
   Check the **HelpCaseLog** sheet to confirm that the case details and a help timestamp have been recorded.

---

## Additional Notes

- **Customizing the Code:**  
  If you need to adjust sheet names, cell references, or add new features, modify the code in `Module1.bas` accordingly.
