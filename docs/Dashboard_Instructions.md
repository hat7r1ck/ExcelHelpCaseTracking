# Dashboard Sheet Instructions

## Overview

The **Dashboard** sheet is designed to provide a visual summary of the help case data logged in the **HelpCaseLog** sheet. It aggregates key performance metrics and displays them in an easy-to-understand format, making it simple for managers and stakeholders to assess your workload and contributions.

## Purpose

The Dashboard serves to:
- **Summarize Performance Metrics:**  
  Show key metrics such as the total number of help cases logged, average response time, and trends over time.
  
- **Visualize Data:**  
  Use charts and graphs to provide a visual representation of your activity.
  
- **Enhance Reporting:**  
  Enable quick insights into periods of high activity, response delays, and overall workload through interactive elements like slicers.

## Data Source

- **Primary Source:**  
  The Dashboard pulls data from the **HelpCaseLog** table/sheet, which contains details of each help case including CaseID, TimeCreated, HelpTimestamp, TimeClosed, and any optional notes.

## Setting Up the Dashboard

### 1. Create PivotTables
- **Insert a PivotTable:**  
  - Select the data range from the **HelpCaseLog** sheet (ideally, the entire table if formatted as an Excel Table).
  - Go to **Insert** → **PivotTable** and choose to place the PivotTable on the Dashboard sheet.
  
- **Key Metrics to Display:**
  - **Total Help Cases:**  
    Count the number of entries.
  - **Average Response Time:**  
    Calculate the average difference between TimeCreated and HelpTimestamp.
  - **Case Trends Over Time:**  
    Group cases by date or month to visualize trends.

### 2. Build Charts and Graphs
- **Bar or Column Charts:**  
  Create charts to show the number of help cases logged per day, week, or month.
- **Line Charts:**  
  Plot average response times over a period to identify trends.
- **Pie Charts:**  
  Use if you have categorical data (e.g., types of cases or statuses) that you want to display.
  
- **Adding Charts:**
  - Once your PivotTable is created, select it and go to **Insert** → choose your desired chart type.
  - Format the chart with clear titles, axis labels, and legends.

### 3. Add Slicers and Filters
- **Insert Slicers:**  
  Slicers enable interactive filtering of data in PivotTables and charts.
  - Click on your PivotTable, then go to **PivotTable Analyze** → **Insert Slicer**.
  - Select relevant fields (e.g., Date, Case Type) to allow for dynamic filtering.
- **Use Timelines:**  
  For date fields, insert a Timeline slicer to filter the data by day, month, or year.

### 4. Design and Layout
- **Organize Dashboard Elements:**  
  Arrange your PivotTables, charts, and slicers in a logical and aesthetically pleasing layout.
- **Consistent Styling:**  
  Use consistent colors, fonts, and labels across all dashboard elements.
- **Annotations:**  
  Consider adding text boxes or labels to provide context or explain specific metrics.

## Customizing the Dashboard

- **Adding New Metrics:**  
  If additional data points are needed (e.g., escalation rates, resolution times), update your PivotTables to include these metrics.
- **Modifying Layout:**  
  Adjust the size and position of your dashboard elements to suit your needs.
- **Updating Data Sources:**  
  If the structure of the **HelpCaseLog** sheet changes (such as adding new columns), ensure that the data range for your PivotTables is updated accordingly. Refresh your PivotTables to load the new data.

## Troubleshooting

- **No Data Displayed:**
  - Verify that the **HelpCaseLog** sheet contains the expected data.
  - Right-click on your PivotTables and select **Refresh**.
  
- **Incorrect Metrics:**
  - Check that the fields and calculations in your PivotTables match the data structure.
  - Ensure that date fields are correctly formatted to allow proper grouping.

- **Formatting Issues:**
  - If charts or slicers aren’t displaying correctly, re-check your source data range and refresh the connections.
  - Adjust the layout manually if any elements overlap or are misaligned.

