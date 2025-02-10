# Excel Help Case Tracking Workbook

## Overview

This project provides an Excel-based solution to track help cases—instances when an analyst reaches out for assistance—that aren’t captured in standard metrics. The workbook automatically imports case data from an external source, logs help events with timestamps, and displays key workload metrics on a dashboard. The solution is designed for low maintenance and minimal manual intervention.

## Features

- **Automatic Data Import:**  
  Uses Power Query to load case data (e.g., from XSOAR) into the *Data_Import* sheet.
  
- **Case Logging:**  
  A simple QuickEntry interface lets you enter a CaseID (along with optional notes and your Owner ID). A VBA macro then retrieves the case details and logs the event in the *CaseLog* sheet with various performance metrics:
  - **MTTP (Mean Time To Pickup):** Calculated as the time from case creation to QuickEntry time.
  - **MTTR (Mean Time To Resolve):** Calculated as the time from case creation to case closure.
  - **Inter-case Gap:** For owner cases only, it shows the gap between the last closed case (for that owner) and the current entry.
  - **Spike Detection:** Flags if multiple cases are created within a short window.
  - **Late Note Status:** Prompts you when a note is required, and updates automatically when the note is addressed.

- **VBA Automation:**  
  The solution leverages VBA to reduce manual data entry and ensure consistency. The main VBA modules are:
  - `src/Module1.bas`: Contains the core automation routines (data import, case logging, dashboard updates, etc.).
  - `src/CaseLog.bas`: Contains the code that automatically updates the "Late Note Status" (changing "NOTE REQUIRED" to "Note provided") when notes are added.

- **Dashboard Visualization (To-do):**  
  A dashboard provides summary metrics (e.g., total help cases, average MTTP/MTTR, and pending note checks).

## Getting Started

1. **Clone or Download the Repository:**  
   Clone the repository or download the ZIP file and extract it.

2. **Open the Workbook:**  
   Open `ExcelHelpCaseTracking.xlsm` in Microsoft Excel (ensure macros are enabled).

3. **Initialize the Workbook:**  
   Run the `InitializeWorkbook` macro (found in `src/Module1.bas`) to create and configure the required sheets:
   - **Data_Import:** Contains raw case data.
   - **CaseLog:** Logs each case along with calculated metrics.
   - **QuickEntry:** For entering a CaseID, optional notes, and your Owner ID.
   - **Dashboard:** Displays a "Last Updated" timestamp.
   - **Log:** Automatically created for audit logging.

4. **Set Up Data Import:**  
   Follow the instructions in [helpful_excel_functions.md](docs/helpful_excel_functions.md) to configure Power Query for importing your case data.

5. **Log a Case:**  
   - Go to the **QuickEntry** sheet.
   - Enter a valid CaseID in cell **B2**, your optional notes in **B3**, and your Owner ID in **B4**.
   - Click the "Add Case" button (or run the `AddHelpCase` macro) to log the case in the *CaseLog* sheet.

6. **Check Late Notes:**  
   Use the "Late Note Checker" button on the **Dashboard** to run the `CheckLateNotes` macro, which will let you know if any cases still require a note.

## Contributing

Contributions, suggestions, and improvements are welcome! Please open an issue or submit a pull request if you have ideas to enhance this project.

