' Advanced Dark Mode Dashboard Module
' This module creates and manages a dark-themed Excel dashboard with automated data refresh and interactive features.
'
' FEATURES:
' - Automated setup of required sheets (Dashboard, CaseLog, Jira, ToDo, Data_Import, QuickEntry, Log).
' - Applies a dark mode theme (dark background, light text) for high contrast.
' - Defines named ranges for key data entry fields on QuickEntry sheet.
' - Dynamic data refresh from Power Query connections and manual input records.
' - Calculation of key performance metrics (case counts, MTTP, MTTR, spike detection).
' - Logging of important events and updates in the Log sheet.
' - Interactive timeline slicer for date filtering (last 14 days by default).
' - Slicers for Case Owner, Category, Status filters.
' - Toggle button support to show/hide dashboard elements.
' - Creates pivot tables and charts with dark theme styling.
' - Performance optimizations (screen updating off during processing, etc.).
' - Macro can be run on demand or on workbook open for automatic refresh.
'
Option Explicit

' Color and style constants for dark theme
Const COLOR_BG As Long = RGB(46, 46, 46)          ' Dark background (dark gray)
Const COLOR_TEXT As Long = RGB(230, 230, 230)     ' Light text color for contrast (near white)
Const COLOR_ACCENT As Long = RGB(0, 120, 215)     ' Accent color (e.g., blue for highlights)
Const COLOR_HIGHLIGHT As Long = RGB(255, 128, 0)  ' Highlight color (orange) for alerts or emphasis

' Main setup procedure: creates sheets, applies theme, defines named ranges
Public Sub SetupDashboard()
    Dim ws As Worksheet
    ' 1. Create required sheets if they don't exist
    Dim sheetNames As Variant
    sheetNames = Array("Dashboard", "CaseLog", "Jira", "ToDo", "Data_Import", "QuickEntry", "Log")
    Dim name As Variant
    For Each name In sheetNames
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(name)
        On Error GoTo 0
        If ws Is Nothing Then
            ' Sheet not found, so add it at end
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            ws.Name = CStr(name)
        End If
        ' Apply dark theme formatting to the sheet
        ApplyDarkTheme ws
        Set ws = Nothing
    Next name
    
    ' (Optional) Create a hidden sheet for pivot tables if not present
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("DashboardPivot")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "DashboardPivot"
        ws.Visible = xlSheetHidden    ' hide the pivot data sheet from view
    End If
    ApplyDarkTheme ws
    Set ws = Nothing
    
    ' 2. Set up QuickEntry sheet labels and named input cells
    SetupQuickEntryForm
    
    ' 3. Prepare CaseLog sheet (ensure headers and convert to table)
    SetupCaseLogTable
    
    ' 4. Prepare Dashboard sheet layout (metrics placeholders, etc.)
    SetupDashboardLayout
    
    ' 5. Log the setup completion
    LogEvent "Dashboard setup completed."
End Sub

' Apply dark mode formatting to an entire worksheet
Private Sub ApplyDarkTheme(targetSheet As Worksheet)
    With targetSheet.Cells
        .Font.Color = COLOR_TEXT
        .Interior.Color = COLOR_BG
    End With
    ' If this is the Dashboard sheet, hide gridlines for a cleaner look
    If targetSheet.Name = "Dashboard" Then
        Dim currSheet As Worksheet: Set currSheet = ActiveSheet
        targetSheet.Activate
        ActiveWindow.DisplayGridlines = False
        currSheet.Activate
    End If
End Sub

' Create a structured table in CaseLog sheet with appropriate headers if not already set
Private Sub SetupCaseLogTable()
    Dim wsLog As Worksheet
    Set wsLog = Worksheets("CaseLog")
    ' If CaseLog has no headers, insert default headers
    If Application.WorksheetFunction.CountA(wsLog.Rows(1)) = 0 Then
        wsLog.Range("A1").Value = "CaseID"
        wsLog.Range("B1").Value = "Owner"
        wsLog.Range("C1").Value = "Category"
        wsLog.Range("D1").Value = "Status"
        wsLog.Range("E1").Value = "TimeCreated"
        wsLog.Range("F1").Value = "AssignedTime"
        wsLog.Range("G1").Value = "ResolvedTime"
        ' (Add more columns if needed for additional metrics)
    End If
    ' Convert the range into an Excel Table (ListObject) for structured referencing
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = wsLog.ListObjects("tblCaseLog")
    On Error GoTo 0
    If tbl Is Nothing Then
        ' Define range up to at least one data row (or a dummy row if no data yet)
        Dim lastCol As Long, lastRow As Long
        lastCol = wsLog.Cells(1, wsLog.Columns.Count).End(xlToLeft).Column
        lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
        If lastRow < 2 Then lastRow = 2  ' ensure at least one blank data row
        wsLog.ListObjects.Add SourceType:=xlSrcRange, _
                              Source:=wsLog.Range(wsLog.Cells(1, 1), wsLog.Cells(lastRow, lastCol)), _
                              XlListObjectHasHeaders:=xlYes
        Set tbl = wsLog.ListObjects(1)
        tbl.Name = "tblCaseLog"
    End If
    ' Format header row for clarity
    wsLog.Rows(1).Font.Bold = True
End Sub

' Set up the QuickEntry sheet with labels and name important input cells
Private Sub SetupQuickEntryForm()
    Dim wsQE As Worksheet
    Set wsQE = Worksheets("QuickEntry")
    ' If QuickEntry is blank, add field labels and blank cells for inputs
    If Application.WorksheetFunction.CountA(wsQE.Cells) = 0 Then
        wsQE.Range("A1").Value = "Case ID:"
        wsQE.Range("A2").Value = "Owner:"
        wsQE.Range("A3").Value = "Category:"
        wsQE.Range("A4").Value = "Status:"
        wsQE.Range("B1:B4").Value = ""  ' reserve B1:B4 for user inputs
    End If
    ' Define named ranges for the input cells (for easy formula/reference)
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="NewCaseID", RefersTo:=wsQE.Range("B1")
    ThisWorkbook.Names.Add Name:="NewOwner", RefersTo:=wsQE.Range("B2")
    ThisWorkbook.Names.Add Name:="NewCategory", RefersTo:=wsQE.Range("B3")
    ThisWorkbook.Names.Add Name:="NewStatus", RefersTo:=wsQE.Range("B4")
    On Error GoTo 0
End Sub

' Set up the Dashboard sheet layout: placeholders for key metrics values
Private Sub SetupDashboardLayout()
    Dim wsDash As Worksheet
    Set wsDash = Worksheets("Dashboard")
    ' Place labels for metrics
    wsDash.Range("A1").Value = "Total Cases (last 2 wks):"
    wsDash.Range("A2").Value = "Average MTTR (hrs):"
    wsDash.Range("A3").Value = "Average MTTP (hrs):"
    wsDash.Range("A4").Value = "Spike Detected:"
    wsDash.Range("A1:A4").Font.Bold = True
    ' Define named ranges for metrics output cells
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="MetricTotalCases", RefersTo:=wsDash.Range("B1")
    ThisWorkbook.Names.Add Name:="MetricAvgMTTR", RefersTo:=wsDash.Range("B2")
    ThisWorkbook.Names.Add Name:="MetricAvgMTTP", RefersTo:=wsDash.Range("B3")
    ThisWorkbook.Names.Add Name:="MetricSpike", RefersTo:=wsDash.Range("B4")
    On Error GoTo 0
    ' Initialize metrics cells
    wsDash.Range("B1:B4").Value = "N/A"
End Sub

' Refresh all data and update the dashboard (pivots, charts, metrics)
Public Sub RefreshDashboard()
    On Error GoTo Cleanup  ' ensure resources are reset on error
    
    ' 1. Optimize performance settings during processing
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Refreshing data..."
    
    ' 2. Refresh all data connections and queries (Power Query, external data)
    ThisWorkbook.RefreshAll
    
    ' 3. Update or create pivot tables and charts for latest data
    UpdatePivotTablesAndCharts
    
    ' 4. Apply default filter on timeline slicer (last 14 days)
    SetDefaultTimeline
    
    ' 5. Recalculate and display key metrics
    CalculateMetrics
    
    ' 6. Finalize
    Application.StatusBar = "Dashboard updated at " & Format(Now, "hh:nn:ss")
    LogEvent "Dashboard refreshed successfully."
    
Cleanup:
    ' Re-enable normal settings
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub

' Create or refresh pivot tables and charts for the dashboard
Private Sub UpdatePivotTablesAndCharts()
    Dim wsPivot As Worksheet: Set wsPivot = Worksheets("DashboardPivot")
    Dim wsDash As Worksheet: Set wsDash = Worksheets("Dashboard")
    Dim pc As PivotCache
    Dim ptCasesByDate As PivotTable, ptByOwner As PivotTable
    Dim ptByCategory As PivotTable, ptByStatus As PivotTable
    
    ' Use CaseLog table as data source for pivot cache
    Dim srcData As String
    srcData = "tblCaseLog"
    
    If wsPivot.PivotTables.Count = 0 Then
        ' No pivot tables yet: create pivot cache and pivot tables
        Set pc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcData)
        ' Pivot 1: Cases by Date (for timeline)
        Set ptCasesByDate = wsPivot.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPivot.Range("A1"), TableName:="PivotCasesByDate")
        With ptCasesByDate
            .PivotFields("TimeCreated").Orientation = xlRowField
            .AddDataField .PivotFields("CaseID"), "CountCases", xlCount
            .PivotFields("TimeCreated").NumberFormat = "yyyy-mm-dd"
            .RowAxisLayout xlOutlineRow
            .ColumnGrand = False: .RowGrand = False
            .NullString = ""
            .TableStyle2 = "PivotStyleDark2"
        End With
        ' Pivot 2: Cases by Owner
        Set ptByOwner = wsPivot.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPivot.Range("A15"), TableName:="PivotByOwner")
        With ptByOwner
            .PivotFields("Owner").Orientation = xlRowField
            .AddDataField .PivotFields("CaseID"), "CountCases", xlCount
            .ColumnGrand = False: .RowGrand = False
            .TableStyle2 = "PivotStyleDark2"
        End With
        ' Pivot 3: Cases by Category
        Set ptByCategory = wsPivot.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPivot.Range("D15"), TableName:="PivotByCategory")
        With ptByCategory
            .PivotFields("Category").Orientation = xlRowField
            .AddDataField .PivotFields("CaseID"), "CountCases", xlCount
            .ColumnGrand = False: .RowGrand = False
            .TableStyle2 = "PivotStyleDark2"
        End With
        ' Pivot 4: Cases by Status
        Set ptByStatus = wsPivot.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPivot.Range("G15"), TableName:="PivotByStatus")
        With ptByStatus
            .PivotFields("Status").Orientation = xlRowField
            .AddDataField .PivotFields("CaseID"), "CountCases", xlCount
            .ColumnGrand = False: .RowGrand = False
            .TableStyle2 = "PivotStyleDark2"
        End With
        ' Create charts for the pivots on the Dashboard
        CreatePivotChart ptByOwner, wsDash, "OwnerChart", xlColumnClustered, Left:=0, Top:=100, Width:=350, Height:=250, title:="Cases by Owner"
        CreatePivotChart ptByCategory, wsDash, "CategoryChart", xlColumnClustered, Left:=360, Top:=100, Width:=350, Height:=250, title:="Cases by Category"
        CreatePivotChart ptCasesByDate, wsDash, "TrendChart", xlLine, Left:=0, Top:=360, Width:=720, Height:=250, title:="Cases Over Time"
        ' (For brevity, no separate chart for Status; can be added similarly if desired)
        ' Create slicers on Dashboard for Owner, Category, Status
        CreatePivotSlicer wsDash, ptByOwner, "Owner", Left:=720, Top:=100, Width:=150, Height:=100
        CreatePivotSlicer wsDash, ptByCategory, "Category", Left:=720, Top:=210, Width:=150, Height:=100
        CreatePivotSlicer wsDash, ptByStatus, "Status", Left:=720, Top:=320, Width:=150, Height:=100
        ' Create timeline slicer for TimeCreated on Dashboard
        CreateTimelineSlicer wsDash, ptCasesByDate, "TimeCreated", Left:=0, Top:=620, Width:=720, Height:=50
        ' Log creation of pivots and charts
        LogEvent "Pivot tables and charts initialized."
    Else
        ' Pivot tables already exist: just refresh pivot cache to update data
        Dim pcache As PivotCache
        For Each pcache In ActiveWorkbook.PivotCaches
            On Error Resume Next
            pcache.Refresh
            On Error GoTo 0
        Next pcache
    End If
End Sub

' Create a pivot chart from a PivotTable and format it for dark theme
Private Sub CreatePivotChart(pivotTbl As PivotTable, targetSheet As Worksheet, chartName As String, chartType As XlChartType, _
                              Left As Double, Top As Double, Width As Double, Height As Double, Optional title As String = "")
    Dim chtObj As ChartObject
    Set chtObj = targetSheet.ChartObjects.Add(Left, Top, Width, Height)
    chtObj.Name = chartName
    With chtObj.Chart
        .SetSourceData Source:=pivotTbl.TableRange2
        .ChartType = chartType
        If title <> "" Then
            .HasTitle = True
            .ChartTitle.Text = title
            .ChartTitle.Font.Color = COLOR_TEXT
        End If
        ' Format chart area and plot area for dark background
        .ChartArea.Format.Fill.ForeColor.RGB = COLOR_BG
        .PlotArea.Format.Fill.ForeColor.RGB = COLOR_BG
        .ChartArea.Format.Line.Visible = msoFalse
        If .SeriesCollection.Count = 1 Then
            .HasLegend = False
        Else
            .Legend.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_TEXT
        End If
        ' Format axes text to light color (if axes exist for this chart type)
        On Error Resume Next
        .Axes(xlCategory).Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_TEXT
        .Axes(xlValue).Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_TEXT
        On Error GoTo 0
    End With
End Sub

' Create a slicer on the dashboard for a given pivot field (non-date)
Private Sub CreatePivotSlicer(targetSheet As Worksheet, pivotTbl As PivotTable, fieldName As String, _
                               Left As Double, Top As Double, Width As Double, Height As Double)
    Dim slicerCache As SlicerCache
    Set slicerCache = ActiveWorkbook.SlicerCaches.Add2(pivotTbl, pivotTbl.PivotFields(fieldName))
    slicerCache.Slicers.Add SlicerDestination:=targetSheet, Left:=Left, Top:=Top, Width:=Width, Height:=Height
    ' Apply a dark slicer style
    slicerCache.Slicers(1).Style = "SlicerStyleDark1"
End Sub

' Create a timeline slicer (for date field) on the dashboard
Private Sub CreateTimelineSlicer(targetSheet As Worksheet, pivotTbl As PivotTable, dateFieldName As String, _
                                  Left As Double, Top As Double, Width As Double, Height As Double)
    Dim slicerCache As SlicerCache
    Set slicerCache = ActiveWorkbook.SlicerCaches.Add2(pivotTbl, dateFieldName, , xlTimeline)
    slicerCache.Slicers.Add SlicerDestination:=targetSheet, Left:=Left, Top:=Top, Width:=Width, Height:=Height
    ' (Default timeline style is applied; can be formatted if needed)
End Sub

' Set timeline slicer to show the last 14 days by default
Private Sub SetDefaultTimeline()
    On Error Resume Next
    Dim sc As SlicerCache
    For Each sc In ActiveWorkbook.SlicerCaches
        If sc.SlicerCacheType = xlTimeline Then
            Dim startDate As Date, endDate As Date
            endDate = Date
            startDate = endDate - 14
            sc.TimelineState.SetFilterDateRange startDate, endDate
        End If
    Next sc
    On Error GoTo 0
End Sub

' Calculate key performance metrics and update the dashboard metrics cells
Private Sub CalculateMetrics()
    Dim wsLog As Worksheet: Set wsLog = Worksheets("CaseLog")
    Dim wsDash As Worksheet: Set wsDash = Worksheets("Dashboard")
    Dim totalCases As Long, countResolved As Long, countAssigned As Long
    Dim sumTTR As Double, sumTTP As Double
    Dim avgMTTR As Double, avgMTTP As Double
    totalCases = 0: countResolved = 0: countAssigned = 0
    Dim cutoffDate As Date
    cutoffDate = Date - 14   ' look-back period of 14 days
    
    ' Iterate through all cases in the CaseLog table
    Dim tbl As ListObject: Set tbl = wsLog.ListObjects("tblCaseLog")
    Dim rw As ListRow
    For Each rw In tbl.ListRows
        Dim created As Variant, assigned As Variant, resolved As Variant
        created = rw.Range.Cells(1, tbl.ListColumns("TimeCreated").Index).Value
        assigned = rw.Range.Cells(1, tbl.ListColumns("AssignedTime").Index).Value
        resolved = rw.Range.Cells(1, tbl.ListColumns("ResolvedTime").Index).Value
        
        If Not IsEmpty(created) Then
            If created >= cutoffDate Then
                totalCases = totalCases + 1
                ' MTTR calculation (if resolved date is available)
                If Not IsEmpty(resolved) Then
                    countResolved = countResolved + 1
                    sumTTR = sumTTR + DateDiff("n", created, resolved)
                End If
                ' MTTP calculation (if assigned/pickup time is available)
                If Not IsEmpty(assigned) Then
                    countAssigned = countAssigned + 1
                    sumTTP = sumTTP + DateDiff("n", created, assigned)
                End If
            End If
        End If
    Next rw
    
    ' Compute averages (convert minutes to hours for MTTR/MTTP)
    If countResolved > 0 Then
        avgMTTR = (sumTTR / countResolved) / 60
    Else
        avgMTTR = 0
    End If
    If countAssigned > 0 Then
        avgMTTP = (sumTTP / countAssigned) / 60
    Else
        avgMTTP = 0
    End If
    
    ' Update dashboard metric cells
    Range("MetricTotalCases").Value = totalCases
    Range("MetricAvgMTTR").Value = Format(avgMTTR, "0.0")
    Range("MetricAvgMTTP").Value = Format(avgMTTP, "0.0")
    
    ' Spike detection: flag if any single day's count > 1.5x average daily count in period
    Dim spikeMsg As String: spikeMsg = "No"
    If totalCases > 0 Then
        Dim avgDaily As Double: avgDaily = totalCases / 14
        Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
        Dim key As Variant, cnt As Long, maxDaily As Long, spikeDay As String
        maxDaily = 0
        ' Count cases per day in the last 14 days
        For Each rw In tbl.ListRows
            If Not IsEmpty(rw.Range.Cells(1, tbl.ListColumns("TimeCreated").Index).Value) Then
                Dim dt As Date
                dt = rw.Range.Cells(1, tbl.ListColumns("TimeCreated").Index).Value
                If dt >= cutoffDate Then
                    key = Format(dt, "yyyy-mm-dd")
                    If dict.Exists(key) Then
                        dict(key) = dict(key) + 1
                    Else
                        dict.Add key, 1
                    End If
                End If
            End If
        Next rw
        ' Determine max daily count
        For Each key In dict.Keys
            cnt = CLng(dict(key))
            If cnt > maxDaily Then
                maxDaily = cnt
                spikeDay = CStr(key)
            End If
        Next key
        If maxDaily > 1.5 * avgDaily And maxDaily > 0 Then
            spikeMsg = "Yes (" & maxDaily & " cases on " & spikeDay & ")"
            LogEvent "Spike detected: " & maxDaily & " cases on " & spikeDay & " (>" & Format(1.5 * avgDaily, "0") & " avg)."
        End If
    End If
    Range("MetricSpike").Value = spikeMsg
    ' Highlight the spike metric if a spike was detected
    If spikeMsg <> "No" Then
        Range("MetricSpike").Font.Color = COLOR_HIGHLIGHT
    Else
        Range("MetricSpike").Font.Color = COLOR_TEXT
    End If
End Sub

' Toggle visibility of a named Shape (chart, slicer, etc.) on the Dashboard (for toggle button functionality)
Public Sub ToggleElement(elementName As String)
    Dim wsDash As Worksheet: Set wsDash = Worksheets("Dashboard")
    Dim shp As Shape
    On Error Resume Next
    Set shp = wsDash.Shapes(elementName)
    On Error GoTo 0
    If Not shp Is Nothing Then
        shp.Visible = Not shp.Visible   ' toggle the shape's visibility
        LogEvent "Toggled visibility of '" & elementName & "' to " & IIf(shp.Visible, "Shown", "Hidden") & "."
    Else
        MsgBox "Dashboard element '" & elementName & "' not found.", vbExclamation
    End If
End Sub

' Log an event with a timestamp in the Log sheet
Private Sub LogEvent(message As String)
    Dim wsLog As Worksheet: Set wsLog = Worksheets("Log")
    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2  ' start at row 2 if log was previously empty
    wsLog.Cells(nextRow, 1).Value = Now
    wsLog.Cells(nextRow, 2).Value = message
    wsLog.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:nn:ss"
    wsLog.Columns(1).AutoFit: wsLog.Columns(2).AutoFit
End Sub

' Automation on Open (Optional): To automatically refresh the dashboard whenever the workbook is opened
'Private Sub Workbook_Open()  
'    Call RefreshDashboard  
'End Sub
