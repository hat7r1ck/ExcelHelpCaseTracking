' Module: DashboardModule.bas
'
Option Explicit

'==============================
' Color constants for modern theme
'==============================
Const LIGHT_BG As Long = 15790320      ' RGB(240,240,240)
Const DARK_TEXT As Long = 4210752       ' RGB(64,64,64)
Const BLUE_ACCENT As Long = 14120960     ' RGB(0,120,215)
Const BLUE_HIGHLIGHT As Long = 13395456  ' RGB(0,102,204)

'==============================
' Sheet names
'==============================
Const SHEET_CASELOG As String = "CaseLog"
Const SHEET_JIRA As String = "Jira"
Const SHEET_TODO As String = "ToDo"
Const SHEET_DASHBOARD As String = "Dashboard"

' Global PivotTable names for each section
Const PT_CASELOG As String = "ptCaseLog"
Const PT_JIRA As String = "ptJira"
Const PT_TODO As String = "ptTodo"

'==============================
' SetupDashboardEnvironment
' Ensures that required sheets exist and applies the theme on the Dashboard.
'==============================
Public Sub SetupDashboardEnvironment()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    Dim requiredSheets As Variant, sName As Variant
    requiredSheets = Array(SHEET_CASELOG, SHEET_JIRA, SHEET_TODO, SHEET_DASHBOARD)
    For Each sName In requiredSheets
        On Error Resume Next
        Set ws = wb.Sheets(sName)
        On Error GoTo 0
        If ws Is Nothing Then
            Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
            ws.Name = sName
        End If
        If sName = SHEET_DASHBOARD Then
            ApplyTheme ws
        End If
    Next sName
    MsgBox "Dashboard environment set up.", vbInformation, "Setup Complete"
End Sub

'==============================
' ApplyTheme
' Applies the modern light theme (light background, dark text, blue accents) to a worksheet.
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
' CreateCaseLogPivot
' Creates a PivotTable from the CaseLog sheet on the Dashboard.
' PivotTable will show case trends by TimeCreated and break down by Owner.
'==============================
Public Sub CreateCaseLogPivot()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsSource As Worksheet: Set wsSource = wb.Sheets(SHEET_CASELOG)
    Dim wsDash As Worksheet: Set wsDash = wb.Sheets(SHEET_DASHBOARD)
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range
    Dim ptCache As PivotCache, pt As PivotTable
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    Set dataRange = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol))
    
    ' Remove existing pivot if it exists (assume location D2)
    On Error Resume Next
    wsDash.PivotTables(PT_CASELOG).TableRange2.Clear
    On Error GoTo 0
    
    Set ptCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        dataRange.Address(ReferenceStyle:=xlR1C1, External:=True))
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsDash.Range("D2"), TableName:=PT_CASELOG)
    
    With pt
        On Error Resume Next
        ' Use TimeCreated as Row field
        .PivotFields("TimeCreated").Orientation = xlRowField
        .PivotFields("TimeCreated").Position = 1
        ' Use Owner as Column field (if available)
        .PivotFields("Owner").Orientation = xlColumnField
        .PivotFields("Owner").Position = 1
        ' Data field: count of CaseID
        .AddDataField .PivotFields("CaseID"), "Case Count", xlCount
        On Error GoTo 0
    End With
    pt.TableStyle2 = "PivotStyleLight16"
End Sub

'==============================
' CreateJiraPivot
' Creates a PivotTable from the Jira sheet on the Dashboard.
' PivotTable will show ticket trends by DateTimeReceived and break down by Confirmation.
'==============================
Public Sub CreateJiraPivot()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsSource As Worksheet: Set wsSource = wb.Sheets(SHEET_JIRA)
    Dim wsDash As Worksheet: Set wsDash = wb.Sheets(SHEET_DASHBOARD)
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range
    Dim ptCache As PivotCache, pt As PivotTable
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    Set dataRange = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol))
    
    ' Remove existing pivot if exists (assume location D20)
    On Error Resume Next
    wsDash.PivotTables(PT_JIRA).TableRange2.Clear
    On Error GoTo 0
    
    Set ptCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        dataRange.Address(ReferenceStyle:=xlR1C1, External:=True))
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsDash.Range("D20"), TableName:=PT_JIRA)
    
    With pt
        On Error Resume Next
        .PivotFields("DateTimeReceived").Orientation = xlRowField
        .PivotFields("DateTimeReceived").Position = 1
        .PivotFields("Confirmation").Orientation = xlColumnField
        .PivotFields("Confirmation").Position = 1
        .AddDataField .PivotFields("Subject"), "Ticket Count", xlCount
        On Error GoTo 0
    End With
    pt.TableStyle2 = "PivotStyleLight16"
End Sub

'==============================
' CreateTodoPivot
' Creates a PivotTable from the ToDo sheet on the Dashboard.
' PivotTable will show task count by Status and break down by Priority.
'==============================
Public Sub CreateTodoPivot()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsSource As Worksheet: Set wsSource = wb.Sheets(SHEET_TODO)
    Dim wsDash As Worksheet: Set wsDash = wb.Sheets(SHEET_DASHBOARD)
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range
    Dim ptCache As PivotCache, pt As PivotTable
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    Set dataRange = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol))
    
    ' Remove existing pivot if exists (assume location D38)
    On Error Resume Next
    wsDash.PivotTables(PT_TODO).TableRange2.Clear
    On Error GoTo 0
    
    Set ptCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        dataRange.Address(ReferenceStyle:=xlR1C1, External:=True))
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsDash.Range("D38"), TableName:=PT_TODO)
    
    With pt
        On Error Resume Next
        .PivotFields("Status").Orientation = xlRowField
        .PivotFields("Status").Position = 1
        .PivotFields("Priority").Orientation = xlColumnField
        .PivotFields("Priority").Position = 1
        .AddDataField .PivotFields("Task"), "Task Count", xlCount
        On Error GoTo 0
    End With
    pt.TableStyle2 = "PivotStyleLight16"
End Sub

'==============================
' CreateCharts
' Creates PivotCharts for each PivotTable.
'==============================
Public Sub CreateCharts()
    Dim wsDash As Worksheet: Set wsDash = ThisWorkbook.Sheets(SHEET_DASHBOARD)
    ' Create CaseLog Trend Chart (line chart) at cell L2
    Call CreatePivotChart(ThisWorkbook.Sheets(SHEET_DASHBOARD).PivotTables(PT_CASELOG), wsDash, "CaseLogChart", xlLine, _
        wsDash.Range("L2").Left, wsDash.Range("L2").Top, 400, 250, "CaseLog Trends")
    ' Create Jira Trend Chart (line chart) at cell L20
    Call CreatePivotChart(ThisWorkbook.Sheets(SHEET_DASHBOARD).PivotTables(PT_JIRA), wsDash, "JiraChart", xlLine, _
        wsDash.Range("L20").Left, wsDash.Range("L20").Top, 400, 250, "Jira Tickets Trends")
    ' Create ToDo Chart (column chart) at cell L38
    Call CreatePivotChart(ThisWorkbook.Sheets(SHEET_DASHBOARD).PivotTables(PT_TODO), wsDash, "TodoChart", xlColumnClustered, _
        wsDash.Range("L38").Left, wsDash.Range("L38").Top, 400, 250, "ToDo Tasks Breakdown")
End Sub

'==============================
' CreatePivotChart
' Creates a PivotChart based on a given PivotTable.
'==============================
Private Sub CreatePivotChart(pt As PivotTable, wsDash As Worksheet, chartName As String, chartType As XlChartType, _
                               chartLeft As Double, chartTop As Double, chartWidth As Double, chartHeight As Double, ChartTitle As String)
    Dim chtObj As ChartObject
    Set chtObj = wsDash.ChartObjects.Add(Left:=chartLeft, Top:=chartTop, Width:=chartWidth, Height:=chartHeight)
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
' Adds interactive slicers for each PivotTable (if the field exists).
'==============================
Public Sub CreateSlicers()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsDash As Worksheet: Set wsDash = wb.Sheets(SHEET_DASHBOARD)
    Dim pt As PivotTable
    Dim sc As SlicerCache, sl As Slicer
    
    ' Delete existing slicers
    Dim scCache As SlicerCache
    For Each scCache In wb.SlicerCaches
        For Each sl In scCache.Slicers
            sl.Delete
        Next sl
    Next scCache
    
    ' For CaseLog pivot: Timeline slicer for TimeCreated and slicer for Owner
    On Error Resume Next
    Set pt = wb.Sheets(SHEET_DASHBOARD).PivotTables(PT_CASELOG)
    Dim tlCache As SlicerCache
    Set tlCache = wb.SlicerCaches.Add2(pt, "TimeCreated", , xlTimeline)
    tlCache.Slicers.Add wsDash, , "Timeline_CaseLog", "Time Created", wsDash.Range("P2")
    Set sc = wb.SlicerCaches.Add2(pt, pt.PivotFields("Owner"))
    sc.Slicers.Add wsDash, , "Slicer_CaseLog_Owner", "Owner", wsDash.Range("P15")
    On Error GoTo 0
    
    ' For Jira pivot: Timeline slicer for DateTimeReceived and slicer for Confirmation
    On Error Resume Next
    Set pt = wb.Sheets(SHEET_DASHBOARD).PivotTables(PT_JIRA)
    Set tlCache = wb.SlicerCaches.Add2(pt, "DateTimeReceived", , xlTimeline)
    tlCache.Slicers.Add wsDash, , "Timeline_Jira", "Received", wsDash.Range("P30")
    Set sc = wb.SlicerCaches.Add2(pt, pt.PivotFields("Confirmation"))
    sc.Slicers.Add wsDash, , "Slicer_Jira_Confirm", "Confirmation", wsDash.Range("P45")
    On Error GoTo 0
    
    ' For ToDo pivot: Slicer for Status and for Priority
    On Error Resume Next
    Set pt = wb.Sheets(SHEET_DASHBOARD).PivotTables(PT_TODO)
    Set sc = wb.SlicerCaches.Add2(pt, pt.PivotFields("Status"))
    sc.Slicers.Add wsDash, , "Slicer_Todo_Status", "Status", wsDash.Range("P60")
    Set sc = wb.SlicerCaches.Add2(pt, pt.PivotFields("Priority"))
    sc.Slicers.Add wsDash, , "Slicer_Todo_Priority", "Priority", wsDash.Range("P75")
    On Error GoTo 0
End Sub

'==============================
' UpdateMetrics
' Calculates and displays key metrics on the Dashboard.
' For CaseLog: total cases and average MTTR.
' For Jira: total tickets.
' For ToDo: total tasks and percentage completed.
' Metrics are placed in designated cells on the Dashboard.
'==============================
Public Sub UpdateMetrics()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsDash As Worksheet: Set wsDash = wb.Sheets(SHEET_DASHBOARD)
    Dim wsCase As Worksheet: Set wsCase = wb.Sheets(SHEET_CASELOG)
    Dim wsJira As Worksheet: Set wsJira = wb.Sheets(SHEET_JIRA)
    Dim wsTodo As Worksheet: Set wsTodo = wb.Sheets(SHEET_TODO)
    Dim lastRow As Long, i As Long
    Dim totalCases As Long, resolvedCases As Long, totalMTTR As Double
    Dim totalJira As Long
    Dim totalTasks As Long, sumPercent As Double
    
    ' --- CaseLog Metrics ---
    lastRow = wsCase.Cells(wsCase.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        totalCases = totalCases + 1
        If IsDate(wsCase.Cells(i, "C").Value) And IsDate(wsCase.Cells(i, "D").Value) Then
            totalMTTR = totalMTTR + ((wsCase.Cells(i, "D").Value - wsCase.Cells(i, "C").Value) * 24)
            resolvedCases = resolvedCases + 1
        End If
    Next i
    wsDash.Range("B2").Value = totalCases
    If resolvedCases > 0 Then
        wsDash.Range("B3").Value = Format(totalMTTR / resolvedCases, "0.0") & " hrs"
    Else
        wsDash.Range("B3").Value = "N/A"
    End If
    
    ' --- Jira Metrics ---
    lastRow = wsJira.Cells(wsJira.Rows.Count, "A").End(xlUp).Row
    totalJira = lastRow - 1  ' subtract header
    wsDash.Range("B4").Value = totalJira
    
    ' --- ToDo Metrics ---
    lastRow = wsTodo.Cells(wsTodo.Rows.Count, "A").End(xlUp).Row
    totalTasks = lastRow - 1
    For i = 2 To lastRow
        ' Assuming % Completed is in column F as a percentage (e.g., 75 for 75%)
        If IsNumeric(wsTodo.Cells(i, "F").Value) Then
            sumPercent = sumPercent + wsTodo.Cells(i, "F").Value
        End If
    Next i
    wsDash.Range("B5").Value = totalTasks
    If totalTasks > 0 Then
        wsDash.Range("B6").Value = Format(sumPercent / totalTasks, "0.0") & "%"
    Else
        wsDash.Range("B6").Value = "N/A"
    End If
End Sub

'==============================
' RefreshDashboard
' Runs all operations: create pivots, charts, slicers, and update metrics.
'==============================
Public Sub RefreshDashboard()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Create/refresh pivot tables for each source
    CreateCaseLogPivot
    CreateJiraPivot
    CreateTodoPivot
    
    ' Create associated charts
    CreateCharts
    
    ' Create slicers
    CreateSlicers
    
    ' Update key metrics
    UpdateMetrics
    
    ' Update dashboard timestamp (assume cell A1 on Dashboard)
    ThisWorkbook.Sheets(SHEET_DASHBOARD).Range("A1").Value = "Last Updated: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Dashboard refreshed.", vbInformation, "Refresh Complete"
End Sub
