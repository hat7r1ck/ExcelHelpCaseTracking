' ============================================
' MODULE: Advanced Dark Mode Dashboard
' This module creates and manages a dark-themed Excel dashboard
' with automation, pivot tables, and slicers for interactive analysis.
' ============================================

Option Explicit

' Color Constants for Dark Mode
Public Const COLOR_BG As Long = &H2E2E2E ' Dark gray background
Public Const COLOR_TEXT As Long = &HE6E6E6 ' Light text
Public Const COLOR_ACCENT As Long = &H0078D7 ' Blue accent
Public Const COLOR_HIGHLIGHT As Long = &HFF8000 ' Orange highlight

' ============================================
' DASHBOARD SETUP - Initializes all sheets, pivots, and visuals
' ============================================
Public Sub SetupDashboard()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array("Dashboard", "CaseLog", "Jira", "ToDo", "Data_Import", "QuickEntry", "Log")

    ' Create required sheets if they don't exist
    For Each ws In ThisWorkbook.Worksheets
        ApplyDarkTheme ws
    Next ws

    ' Ensure DashboardPivot exists (for PivotTables)
    If Not SheetExists("DashboardPivot") Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "DashboardPivot"
        ws.Visible = xlSheetHidden
    End If

    ' Setup dashboard elements
    SetupQuickEntryForm
    SetupCaseLogTable
    SetupDashboardLayout
    SetupPivotTables

    LogEvent "Dashboard Setup Completed."
    MsgBox "Dark Mode Dashboard Ready!", vbInformation, "Setup Complete"
End Sub

' ============================================
' APPLY DARK MODE TO SHEETS
' ============================================
Private Sub ApplyDarkTheme(targetSheet As Worksheet)
    With targetSheet.Cells
        .Font.Color = COLOR_TEXT
        .Interior.Color = COLOR_BG
    End With
    ' Hide gridlines for cleaner UI
    If targetSheet.Name = "Dashboard" Then
        ActiveWindow.DisplayGridlines = False
    End If
End Sub

' ============================================
' PIVOT TABLES & DATA VISUALIZATION
' ============================================
Private Sub SetupPivotTables()
    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim ptCache As PivotCache, pt As PivotTable
    Dim dataRange As Range

    If Not SheetExists("CaseLog") Then Exit Sub
    Set wsData = ThisWorkbook.Sheets("CaseLog")

    On Error Resume Next
    Set dataRange = wsData.Range("A1").CurrentRegion
    On Error GoTo 0
    If dataRange Is Nothing Then Exit Sub

    Set wsPivot = ThisWorkbook.Sheets("DashboardPivot")
    Set ptCache = ThisWorkbook.PivotCaches.Create(xlDatabase, dataRange)

    ' Remove old pivot table if it exists
    On Error Resume Next
    wsPivot.PivotTables("Pivot_CaseLog").TableRange2.Clear
    On Error GoTo 0

    ' Create PivotTable
    Set pt = wsPivot.PivotTables.Add(PivotCache:=ptCache, TableDestination:=wsPivot.Range("A1"), TableName:="Pivot_CaseLog")

    ' Configure Pivot Fields
    With pt
        .PivotFields("TimeCreated").Orientation = xlRowField
        .AddDataField .PivotFields("CaseID"), "CountCases", xlCount
        .PivotFields("TimeCreated").NumberFormat = "yyyy-mm-dd"
        .RowAxisLayout xlOutlineRow
        .ColumnGrand = False: .RowGrand = False
        .NullString = ""
    End With

    pt.TableStyle2 = "PivotStyleDark1"
    LogEvent "PivotTables Initialized."
End Sub

' ============================================
' INTERACTIVE ELEMENTS: SLICERS & TIMELINE
' ============================================
Private Sub SetupSlicers()
    Dim ws As Worksheet, pt As PivotTable
    Dim slicerCache As SlicerCache, slicer As Slicer

    Set ws = ThisWorkbook.Sheets("Dashboard")
    Set pt = ws.PivotTables("Pivot_CaseLog")

    ' Remove old slicers
    On Error Resume Next
    ThisWorkbook.SlicerCaches("Owner").Delete
    On Error GoTo 0

    ' Create slicer
    Set slicerCache = ThisWorkbook.SlicerCaches.Add(pt, pt.PivotFields("Owner"))
    Set slicer = slicerCache.Slicers.Add(ws, , "Owner", "Owner", 100, 50, 200, 100)

    slicer.Shape.Fill.ForeColor.RGB = COLOR_ACCENT
End Sub

' ============================================
' DASHBOARD AUTOMATION
' ============================================
Public Sub RefreshDashboard()
    On Error GoTo Cleanup

    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Refreshing..."

    ' Refresh Data & Pivots
    ThisWorkbook.RefreshAll
    SetupPivotTables

    ' Apply Date Filter
    SetDefaultDateRange

    ' Recalculate Metrics
    CalculateMetrics

    ' Finalize
    Application.StatusBar = "Dashboard Updated"
    LogEvent "Dashboard Refreshed."

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub

' ============================================
' DATE FILTERING & DEFAULT TIME RANGE
' ============================================
Private Sub SetDefaultDateRange()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")

    ws.Range("EndDate").Value = Date
    ws.Range("StartDate").Value = Date - 13

    ApplyDateFilter
End Sub

Private Sub ApplyDateFilter()
    Dim ws As Worksheet, pt As PivotTable, pf As PivotField
    Dim startDate As Date, endDate As Date

    Set ws = ThisWorkbook.Sheets("Dashboard")

    startDate = ws.Range("StartDate").Value
    endDate = ws.Range("EndDate").Value

    For Each pt In ws.PivotTables
        On Error Resume Next
        Set pf = pt.PivotFields("TimeCreated")
        On Error GoTo 0
        If Not pf Is Nothing Then
            pf.ClearAllFilters
            pf.PivotFilters.Add2 Type:=xlDateBetween, Value1:=startDate, Value2:=endDate
        End If
    Next pt
End Sub

' ============================================
' LOGGING & UTILITY FUNCTIONS
' ============================================
Private Sub LogEvent(message As String)
    Dim ws As Worksheet, nextRow As Long
    Set ws = ThisWorkbook.Sheets("Log")
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = message
    ws.Columns.AutoFit
End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
