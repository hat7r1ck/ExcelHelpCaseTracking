' Module: DashboardModule.bas

Option Explicit

'==============================
' Color constants for modern theme
'==============================
Const LIGHT_BG As Long = 15790320      ' RGB(240,240,240) - light grey background
Const DARK_TEXT As Long = 4210752       ' RGB(64,64,64) - dark text
Const BLUE_ACCENT As Long = 14120960     ' RGB(0,120,215) - blue accent
Const BLUE_HIGHLIGHT As Long = 13395456  ' RGB(0,102,204) - blue highlight

'==============================
' Sheet names
'==============================
Const SHEET_CASELOG As String = "CaseLog"
Const SHEET_JIRA As String = "Jira"
Const SHEET_TODO As String = "ToDo"
Const SHEET_DASHBOARD As String = "Dashboard"
Const SHEET_DATA As String = "DashboardData"  ' Hidden consolidated data sheet

' Global pivot table name
Const PT_NAME As String = "ptDashboard"

'==============================
' SetupDashboardEnvironment
' Ensures that required sheets exist, applies theme to Dashboard, and hides the consolidated data sheet.
'==============================
Public Sub SetupDashboardEnvironment()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim sName As Variant
    Dim i As Long
    
    sheetNames = Array(SHEET_CASELOG, SHEET_JIRA, SHEET_TODO, SHEET_DASHBOARD, SHEET_DATA)
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = wb.Sheets(sheetNames(i))
        On Error GoTo 0
        If ws Is Nothing Then
            Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
            ws.Name = sheetNames(i)
        End If
        
        ' Apply the theme on the Dashboard sheet only
        If ws.Name = SHEET_DASHBOARD Then
            ApplyTheme ws
        End If
        
        ' Hide the consolidated data sheet
        If ws.Name = SHEET_DATA Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next i
    
    MsgBox "Dashboard environment set up.", vbInformation, "Setup Complete"
End Sub

'==============================
' ApplyTheme
' Applies a modern light theme (light background, dark text, blue accents) to the given worksheet.
'==============================
Private Sub ApplyTheme(ws As Worksheet)
    With ws.Cells
        .Interior.Color = LIGHT_BG
        .Font.Color = DARK_TEXT
        .Font.Name = "Calibri"
    End With
    ws.Tab.Color = BLUE_ACCENT
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

'==============================
' ConsolidateData
' Consolidates data from "CaseLog", "Jira", and "ToDo" sheets into the "DashboardData" sheet.
' Assumes headers are in row 1.
'==============================
Public Sub ConsolidateData()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsData As Worksheet: Set wsData = wb.Sheets(SHEET_DATA)
    Dim wsSource As Worksheet
    Dim sourceSheets As Variant
    Dim lastRowData As Long, lastRowSource As Long
    Dim headerRow As Range
    Dim shtName As Variant
    
    sourceSheets = Array(SHEET_CASELOG, SHEET_JIRA, SHEET_TODO)
    
    ' Clear existing data
    wsData.Cells.Clear
    
    ' Copy header row from the first source sheet
    Set wsSource = wb.Sheets(sourceSheets(0))
    Set headerRow = wsSource.Range("A1").CurrentRegion.Rows(1)
    headerRow.Copy Destination:=wsData.Range("A1")
    lastRowData = 1  ' Start after header
    
    ' Append data from each source sheet
    For Each shtName In sourceSheets
        Set wsSource = wb.Sheets(shtName)
        lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
        If lastRowSource > 1 Then
            wsSource.Range("A2", wsSource.Cells(lastRowSource, headerRow.Columns.Count)).Copy _
                Destination:=wsData.Cells(lastRowData + 1, 1)
            lastRowData = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
        End If
    Next shtName
End Sub

'==============================
' CreatePivotDashboard
' Creates a PivotTable (ptDashboard) on the "Dashboard" sheet based on consolidated data,
' and creates a PivotChart (line chart) to display case trends over time.
'==============================
Public Sub CreatePivotDashboard()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsDash As Worksheet: Set wsDash = wb.Sheets(SHEET_DASHBOARD)
    Dim wsData As Worksheet: Set wsData = wb.Sheets(SHEET_DATA)
    Dim ptCache As PivotCache, pt As PivotTable
    Dim dataRange As Range
    Dim lastRow As Long, lastCol As Long
    
    ' Define the data range in DashboardData
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Set dataRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))
    
    ' Remove existing pivot table on Dashboard (if any) starting at cell D10
    Dim ptObj As PivotTable
    On Error Resume Next
    Set ptObj = wsDash.PivotTables(PT_NAME)
    On Error GoTo 0
    If Not ptObj Is Nothing Then
        ptObj.TableRange2.Clear
    End If
    
    ' Create a new pivot cache and pivot table
    Set ptCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        dataRange.Address(ReferenceStyle:=xlR1C1, External:=True))
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsDash.Range("D10"), TableName:=PT_NAME)
    
    ' Configure pivot fields: Use "TimeCreated" in Rows, "Status" in Columns, count of "CaseID" as data.
    With pt
        On Error Resume Next
        .PivotFields("TimeCreated").Orientation = xlRowField
        .PivotFields("TimeCreated").Position = 1
        .PivotFields("Status").Orientation = xlColumnField
        .PivotFields("Status").Position = 1
        .AddDataField .PivotFields("CaseID"), "Case Count", xlCount
        On Error GoTo 0
    End With
    
    ' Apply a clean pivot table style
    pt.TableStyle2 = "PivotStyleLight16"
    
    ' Create a PivotChart (line chart) for case trends over time, placed at cell L10.
    Call CreatePivotChart(pt, wsDash, "TrendChart", xlLine, _
        wsDash.Range("L10").Left, wsDash.Range("L10").Top, 400, 250, "Cases Over Time")
End Sub

'==============================
' CreatePivotChart
' Creates a PivotChart based on a given PivotTable.
' Parameters: pt (PivotTable), wsDash (Dashboard sheet), chartName, chartType,
'             DLeft, DTop, DWidth, DHeight (position/dimensions), ChartTitle.
'==============================
Private Sub CreatePivotChart(pt As PivotTable, wsDash As Worksheet, chartName As String, chartType As XlChartType, _
                               DLeft As Double, DTop As Double, DWidth As Double, DHeight As Double, ChartTitle As String)
    Dim chtObj As ChartObject
    Set chtObj = wsDash.ChartObjects.Add(Left:=DLeft, Top:=DTop, Width:=DWidth, Height:=DHeight)
    chtObj.Name = chartName
    With chtObj.Chart
        .SetSourceData Source:=pt.TableRange2
        .ChartType = chartType
        .HasTitle = True
        .ChartTitle.Text = ChartTitle
        .ChartTitle.Font.Color = BLUE_ACCENT
        .ChartArea.Format.Fill.ForeColor.RGB = LIGHT_BG
        .PlotArea.Format.Fill.ForeColor.RGB = LIGHT_BG
        .Axes(xlCategory).TickLabels.Font.Color = DARK_TEXT
        .Axes(xlValue).TickLabels.Font.Color = DARK_TEXT
    End With
End Sub

'==============================
' CreateSlicers
' Adds interactive slicers (including a Timeline slicer) to the Dashboard.
' Slicers for TimeCreated (Timeline), Status, and Owner.
'==============================
Public Sub CreateSlicers()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim pt As PivotTable: Set pt = wb.Sheets(SHEET_DASHBOARD).PivotTables(PT_NAME)
    Dim wsDash As Worksheet: Set wsDash = wb.Sheets(SHEET_DASHBOARD)
    Dim sc As SlicerCache, sl As Slicer
    
    ' Delete existing slicers
    Dim scCache As SlicerCache
    For Each scCache In wb.SlicerCaches
        For Each sl In scCache.Slicers
            sl.Delete
        Next sl
    Next scCache
    
    ' Create a Timeline slicer for TimeCreated
    On Error Resume Next
    Dim tlSlicerCache As SlicerCache
    Set tlSlicerCache = wb.SlicerCaches.Add2(pt, "TimeCreated", , xlTimeline)
    tlSlicerCache.Slicers.Add wsDash, , "Timeline_TimeCreated", "Time Created", wsDash.Range("D35")
    On Error GoTo 0
    
    ' Create slicer for Status
    Set sc = wb.SlicerCaches.Add2(pt, pt.PivotFields("Status"))
    sc.Slicers.Add wsDash, , "Slicer_Status", "Status", wsDash.Range("J10")
    
    ' Create slicer for Owner (if the field exists)
    On Error Resume Next
    Set sc = wb.SlicerCaches.Add2(pt, pt.PivotFields("Owner"))
    sc.Slicers.Add wsDash, , "Slicer_Owner", "Owner", wsDash.Range("J25")
    On Error GoTo 0
End Sub

'==============================
' UpdateKeyMetrics
' Calculates key metrics from the consolidated data in DashboardData and writes them to Dashboard cells.
' Expected Dashboard cells: B2 (Total Cases), B3 (Open/Closed), B4 (MTTR), B5 (Spike Alert).
'==============================
Public Sub UpdateKeyMetrics()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsDash As Worksheet: Set wsDash = wb.Sheets(SHEET_DASHBOARD)
    Dim wsData As Worksheet: Set wsData = wb.Sheets(SHEET_DATA)
    Dim lastRow As Long, i As Long
    Dim totalCases As Long, openCases As Long, closedCases As Long
    Dim totalResolutionTime As Double, resolvedCount As Long
    Dim dtOpen As Date, dtClosed As Date
    Dim statusVal As String
    
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        totalCases = totalCases + 1
        statusVal = CStr(wsData.Cells(i, "E").Value) ' Assuming "Status" is in column E
        If LCase(statusVal) = "closed" Then
            closedCases = closedCases + 1
            If IsDate(wsData.Cells(i, "C").Value) And IsDate(wsData.Cells(i, "D").Value) Then
                dtOpen = wsData.Cells(i, "C").Value
                dtClosed = wsData.Cells(i, "D").Value
                totalResolutionTime = totalResolutionTime + (dtClosed - dtOpen)
                resolvedCount = resolvedCount + 1
            End If
        Else
            openCases = openCases + 1
        End If
    Next i
    
    wsDash.Range("B2").Value = totalCases
    wsDash.Range("B3").Value = openCases & " Open / " & closedCases & " Closed"
    If resolvedCount > 0 Then
        ' MTTR in hours = average days * 24
        wsDash.Range("B4").Value = Format((totalResolutionTime / resolvedCount) * 24, "0.0") & " hrs"
    Else
        wsDash.Range("B4").Value = "N/A"
    End If
    
    ' Spike detection: if today's case count > 2x average of previous 7 days.
    Dim caseCounts As Object: Set caseCounts = CreateObject("Scripting.Dictionary")
    Dim currentDate As Date, dtKey As Date, countToday As Long, sumPrev As Long, avgPrev As Double
    currentDate = VBA.Int(Date)
    For i = 2 To lastRow
        If IsDate(wsData.Cells(i, "C").Value) Then
            dtKey = VBA.Int(wsData.Cells(i, "C").Value)
            If caseCounts.Exists(dtKey) Then
                caseCounts(dtKey) = caseCounts(dtKey) + 1
            Else
                caseCounts.Add dtKey, 1
            End If
        End If
    Next i
    countToday = 0: sumPrev = 0
    If caseCounts.Exists(currentDate) Then countToday = caseCounts(currentDate)
    Dim d As Long, countPrev As Long
    countPrev = 0
    For d = 1 To 7
        dtKey = currentDate - d
        If caseCounts.Exists(dtKey) Then
            sumPrev = sumPrev + caseCounts(dtKey)
            countPrev = countPrev + 1
        End If
    Next d
    If countPrev > 0 Then
        avgPrev = sumPrev / countPrev
    Else
        avgPrev = 0
    End If
    If avgPrev > 0 And countToday > 2 * avgPrev Then
        wsDash.Range("B5").Value = "YES"
        wsDash.Range("B5").Font.Color = BLUE_HIGHLIGHT
    Else
        wsDash.Range("B5").Value = "NO"
        wsDash.Range("B5").Font.Color = DARK_TEXT
    End If
End Sub

'==============================
' RefreshDashboard
' Consolidates data, refreshes the pivot table/chart, re-creates slicers, and updates key metrics.
'==============================
Public Sub RefreshDashboard()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 1. Consolidate data from source sheets into DashboardData
    ConsolidateData
    
    ' 2. Create/Refresh PivotTable and PivotChart on Dashboard
    CreatePivotDashboard
    
    ' 3. Create slicers for interactivity
    CreateSlicers
    
    ' 4. Update key metrics on Dashboard
    UpdateKeyMetrics
    
    ' Update dashboard timestamp (assume cell B1 on Dashboard)
    ThisWorkbook.Sheets(SHEET_DASHBOARD).Range("B1").Value = "Last Updated: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Dashboard refreshed.", vbInformation, "Refresh Complete"
End Sub
