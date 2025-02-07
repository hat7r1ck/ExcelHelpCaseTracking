# Excel Help Case Tracking Workbook

## Overview

This project provides an Excel-based solution to track help cases—instances when an analyst reaches out for assistance. It was created to solve tracking issues that aren’t captured in standard metrics. The workbook automatically imports case data from an external source, logs help events with a timestamp, and optionally displays workload metrics on a dashboard.

## Features

- **Automatic Data Import:**  
  Uses Power Query to load case data (e.g., from XSOAR) into the Data_Import sheet.
  
- **Help Case Logging:**  
  Provides a simple QuickEntry interface where you enter a CaseID. A VBA macro then looks up the case details and logs a help timestamp in the HelpCaseLog sheet.
  
- **VBA Automation:**  
  The solution leverages VBA (see `src/Module1.bas`) to reduce manual data entry and ensure consistency.
  
- **Dashboard Visualization (Optional):**  
  Includes a dashboard (or can be easily added) to display metrics like total help cases, average response times, and workload trends.
  
- **Low Maintenance:**  
  Designed to work with minimal manual intervention—just update your data source and log cases via the QuickEntry sheet.

## Getting Started

1. **Clone or Download the Repository:**  
   Clone the repository or download the ZIP file and extract it.

2. **Open the Workbook:**  
   Open `ExcelHelpCaseTracking.xlsm` in Microsoft Excel (make sure to enable macros).

3. **Set Up Data Import:**  
   Follow the instructions in [docs/Data_Import_Instructions.md](docs/Data_Import_Instructions.md) to configure Power Query for importing your case data.

4. **Log a Help Case:**  
   - Go to the **QuickEntry** sheet.
   - Enter a valid CaseID in cell B2.
   - Click the "Add Help Case" button to trigger the VBA macro, which logs the case details and a help timestamp in the **HelpCaseLog** sheet.
   
5. **Review Your Data:**  
   - Check the **HelpCaseLog** sheet for the updated log.
   - Use the **Dashboard** sheet (if configured) to view summary metrics.

## Detailed Documentation

For detailed instructions on setting up and using each part of the workbook, please see the following files in the **docs/** folder:

- **Data_Import Sheet:**  
  [docs/Data_Import_Instructions.md](docs/Data_Import_Instructions.md) – Explains how to set up Power Query and format the imported data.

- **HelpCaseLog Sheet:**  
  [docs/HelpCaseLog_Instructions.md](docs/HelpCaseLog_Instructions.md) – Provides guidelines on how to configure the help case log table.

- **QuickEntry Sheet & VBA Macro:**  
  [docs/QuickEntry_Instructions.md](docs/QuickEntry_Instructions.md) – Describes the QuickEntry interface and how the VBA macro (in `src/Module1.bas`) operates.

- **Dashboard Sheet (Optional):**  
  [docs/Dashboard_Instructions.md](docs/Dashboard_Instructions.md) – Offers tips on setting up PivotTables/charts to visualize your metrics.

## Contributing

Contributions, suggestions, and improvements are welcome! Please open an issue or submit a pull request if you have ideas to enhance this project.
