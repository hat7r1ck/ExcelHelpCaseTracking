Option Explicit

' Original theme color constants (pre-calculated)
Const COLOR_BG As Long = 3026478       ' Equivalent to RGB(46,46,46)
Const COLOR_TEXT As Long = 15158550      ' Equivalent to RGB(230,230,230)
Const COLOR_ACCENT As Long = 14120960    ' Equivalent to RGB(0,120,215)
Const COLOR_HIGHLIGHT As Long = 33023    ' Equivalent to RGB(255,128,0)

'***************************************************************
' Advanced Dark Mode Dashboard Module (Original Theme)
' This module creates and manages a dashboard using your original
' theme colors.
'***************************************************************

Public Sub SetupDashboard()
    Dim ws As Worksheet
    Dim sheetNames As Variant, nm As Variant
    Dim wsPivot As Worksheet
    Dim found As Boolean
    Dim sh As Worksheet
    
    ' Check if workbook structure is protected
    If ThisWorkbook.ProtectStructure Then
        MsgBox "Workbook structure is protected. Please unprotect it and try again.", vbCritical
        Exit Sub
    End If
    
    ' List of required sheets
    sheetNames = Array("Dashboard", "CaseLog", "Jira", "ToDo", "Data_Import", "QuickEntry", "Log")
    For Each nm In sheetNames
        found = False
        For Each sh In ThisWorkbook.Worksheets
            If sh.Name = CStr(nm) Then
                Set ws = sh
                found = True
                Exit For
            End If
        Next sh
        If Not found Then
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            On Error Resume Next
            ws.Name = CStr(nm)
            On Error GoTo 0
        Else
            ws.Visible = xlSheetVisible
        End If
        ApplyDarkTheme ws
        Set ws = Nothing
    Next nm
    
    ' Create or find the hidden pivot sheet "DashboardPivot"
    Set wsPivot = Nothing
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name = "DashboardPivot" Then
            Set wsPivot = sh
            Exit For
        End If
    Next sh
    If wsPivot Is Nothing Then
        Set wsPivot = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error Resume Next
        wsPivot.Name = "DashboardPivot"
        On Error GoTo 0
    End If
    wsPivot.Visible = xlSheetHidden
    ApplyDarkTheme wsPivot
    
    ' Set up additional sheets/forms/layouts
    SetupQuickEntryForm
    SetupCaseLogTable
    SetupDashboardLayout
    LogEvent "Dashboard setup completed."
End Sub

Private Sub ApplyDarkTheme(targetSheet As Worksheet)
    With targetSheet.Cells
        .Interior.Color = COLOR_BG
        .Font.Color = COLOR_TEXT
    End With
    ' For the Dashboard sheet, hide gridlines for a cleaner look.
    If targetSheet.Name = "Dashboard" Then
        targetSheet.Activate
        ActiveWindow.DisplayGridlines = False
    End If
End Sub

Private Sub SetupCaseLogTable()
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Worksheets("CaseLog")
    
    ' Add headers if row 1 is empty
    If Application.WorksheetFunction.CountA(wsLog.Rows(1)) = 0 Then
        wsLog.Range("A1:G1").Value = Array("CaseID", "Owner", "Category", "Status", "TimeCreated", "AssignedTime", "ResolvedTime")
    End If
    
    Dim tbl As ListObject
    ' Use existing table if available
    If wsLog.ListObjects.Count > 0 Then
        Set tbl = wsLog.ListObjects(1)
        If tbl.Name <> "tblCaseLog" Then
            On Error Resume Next
            tbl.Name = "tblCaseLog"
            On Error GoTo 0
        End If
    Else
        Dim lastCol As Long, lastRow As Long
        lastCol = wsLog.Cells(1, wsLog.Columns.Count).End(xlToLeft).Column
        lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
        If lastRow < 2 Then lastRow = 2
        Set tbl = wsLog.ListObjects.Add(SourceType:=xlSrcRange, _
                    Source:=wsLog.Range(wsLog.Cells(1, 1), wsLog.Cells(lastRow, lastCol)), _
                    XlListObjectHasHeaders:=xlYes)
        tbl.Name = "tblCaseLog"
    End If
    wsLog.Rows(1).Font.Bold = True
End Sub

Private Sub SetupQuickEntryForm()
    Dim wsQE As Worksheet
    Set wsQE = ThisWorkbook.Worksheets("QuickEntry")
    If Application.WorksheetFunction.CountA(wsQE.Cells) = 0 Then
        wsQE.Range("A1:A4").Value = Application.WorksheetFunction.Transpose(Array("Case ID:", "Owner:", "Category:", "Status:"))
        wsQE.Range("B1:B4").ClearContents
    End If
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="NewCaseID", RefersTo:=wsQE.Range("B1")
    ThisWorkbook.Names.Add Name:="NewOwner", RefersTo:=wsQE.Range("B2")
    ThisWorkbook.Names.Add Name:="NewCategory", RefersTo:=wsQE.Range("B3")
    ThisWorkbook.Names.Add Name:="NewStatus", RefersTo:=wsQE.Range("B4")
    On Error GoTo 0
End Sub

Private Sub SetupDashboardLayout()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    wsDash.Range("A1:A4").Font.Bold = True
    wsDash.Range("A1").Value = "Total Cases (last 2 wks):"
    wsDash.Range("A2").Value = "Average MTTR (hrs):"
    wsDash.Range("A3").Value = "Average MTTP (hrs):"
    wsDash.Range("A4").Value = "Spike Detected:"
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="MetricTotalCases", RefersTo:=wsDash.Range("B1")
    ThisWorkbook.Names.Add Name:="MetricAvgMTTR", RefersTo:=wsDash.Range("B2")
    ThisWorkbook.Names.Add Name:="MetricAvgMTTP", RefersTo:=wsDash.Range("B3")
    ThisWorkbook.Names.Add Name:="MetricSpike", RefersTo:=wsDash.Range("B4")
    On Error GoTo 0
    wsDash.Range("B1:B4").Value = "N/A"
End Sub

Public Sub RefreshDashboard()
    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Refreshing data..."
    
    ThisWorkbook.RefreshAll
    
    UpdatePivotTablesAndCharts
    SetDefaultTimeline
    CalculateMetrics
    
    Application.StatusBar = "Dashboard updated at " & Format(Now, "hh:mm:ss")
    LogEvent "Dashboard refreshed successfully."
    
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub

Private Sub UpdatePivotTablesAndCharts()
    Dim wsPivot As Worksheet, wsDash As Worksheet
    Dim pc As PivotCache
    Dim ptCasesByDate As PivotTable, ptByOwner As PivotTable
    Dim ptByCategory As PivotTable, ptByStatus As PivotTable
    
    Set wsPivot = ThisWorkbook.Worksheets("DashboardPivot")
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    
    Dim srcData As String
    srcData = "tblCaseLog"
    
    If wsPivot.PivotTables.Count = 0 Then
        Set pc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcData)
        
        ' Pivot for Cases by Date
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
        
        ' Pivot for Cases by Owner
        Set ptByOwner = wsPivot.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPivot.Range("A15"), TableName:="PivotByOwner")
        With ptByOwner
            .PivotFields("Owner").Orientation = xlRowField
            .AddDataField .PivotFields("CaseID"), "CountCases", xlCount
            .ColumnGrand = False: .RowGrand = False
            .TableStyle2 = "PivotStyleDark2"
        End With
        
        ' Pivot for Cases by Category
        Set ptByCategory = wsPivot.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPivot.Range("D15"), TableName:="PivotByCategory")
        With ptByCategory
            .PivotFields("Category").Orientation = xlRowField
            .AddDataField .PivotFields("CaseID"), "CountCases", xlCount
            .ColumnGrand = False: .RowGrand = False
            .TableStyle2 = "PivotStyleDark2"
        End With
        
        ' Pivot for Cases by Status
        Set ptByStatus = wsPivot.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPivot.Range("G15"), TableName:="PivotByStatus")
        With ptByStatus
            .PivotFields("Status").Orientation = xlRowField
            .AddDataField .PivotFields("CaseID"), "CountCases", xlCount
            .ColumnGrand = False: .RowGrand = False
            .TableStyle2 = "PivotStyleDark2"
        End With
        
        ' Create charts on Dashboard
        CreatePivotChart ptByOwner, wsDash, "OwnerChart", xlColumnClustered, 0, 100, 350, 250, "Cases by Owner"
        CreatePivotChart ptByCategory, wsDash, "CategoryChart", xlColumnClustered, 360, 100, 350, 250, "Cases by Category"
        CreatePivotChart ptCasesByDate, wsDash, "TrendChart", xlLine, 0, 360, 720, 250, "Cases Over Time"
        
        ' Create slicers and timeline
        CreatePivotSlicer wsDash, ptByOwner, "Owner", 720, 100, 150, 100
        CreatePivotSlicer wsDash, ptByCategory, "Category", 720, 210, 150, 100
        CreatePivotSlicer wsDash, ptByStatus, "Status", 720, 320, 150, 100
        CreateTimelineSlicer wsDash, ptCasesByDate, "TimeCreated", 0, 620, 720, 50
        
        LogEvent "Pivot tables and charts initialized."
    Else
        Dim pcache As PivotCache
        For Each pcache In ActiveWorkbook.PivotCaches
            On Error Resume Next
            pcache.Refresh
            On Error GoTo 0
        Next pcache
    End If
End Sub

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
        .ChartArea.Format.Fill.ForeColor.RGB = COLOR_BG
        .PlotArea.Format.Fill.ForeColor.RGB = COLOR_BG
        .ChartArea.Format.Line.Visible = msoFalse
        On Error Resume Next
        .Axes(xlCategory).Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_TEXT
        .Axes(xlValue).Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_TEXT
        On Error GoTo 0
    End With
End Sub

Private Sub CreatePivotSlicer(targetSheet As Worksheet, pivotTbl As PivotTable, fieldName As String, _
                               Left As Double, Top As Double, Width As Double, Height As Double)
    Dim slicerCache As SlicerCache
    Set slicerCache = ActiveWorkbook.SlicerCaches.Add2(pivotTbl, pivotTbl.PivotFields(fieldName))
    slicerCache.Slicers.Add SlicerDestination:=targetSheet, Left:=Left, Top:=Top, Width:=Width, Height:=Height
    slicerCache.Slicers(1).Style = "SlicerStyleDark1"
End Sub

Private Sub CreateTimelineSlicer(targetSheet As Worksheet, pivotTbl As PivotTable, dateFieldName As String, _
                                  Left As Double, Top As Double, Width As Double, Height As Double)
    Dim slicerCache As SlicerCache
    Set slicerCache = ActiveWorkbook.SlicerCaches.Add2(pivotTbl, dateFieldName, , xlTimeline)
    slicerCache.Slicers.Add SlicerDestination:=targetSheet, Left:=Left, Top:=Top, Width:=Width, Height:=Height
End Sub

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

Private Sub CalculateMetrics()
    Dim wsLog As Worksheet, wsDash As Worksheet
    Set wsLog = Worksheets("CaseLog")
    Set wsDash = Worksheets("Dashboard")
    Dim totalCases As Long, countResolved As Long, countAssigned As Long
    Dim sumTTR As Double, sumTTP As Double
    Dim avgMTTR As Double, avgMTTP As Double
    totalCases = 0: countResolved = 0: countAssigned = 0
    Dim cutoffDate As Date: cutoffDate = Date - 14
    
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
                If Not IsEmpty(resolved) Then
                    countResolved = countResolved + 1
                    sumTTR = sumTTR + DateDiff("n", created, resolved)
                End If
                If Not IsEmpty(assigned) Then
                    countAssigned = countAssigned + 1
                    sumTTP = sumTTP + DateDiff("n", created, assigned)
                End If
            End If
        End If
    Next rw
    
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
    
    Range("MetricTotalCases").Value = totalCases
    Range("MetricAvgMTTR").Value = Format(avgMTTR, "0.0")
    Range("MetricAvgMTTP").Value = Format(avgMTTP, "0.0")
    
    Dim spikeMsg As String: spikeMsg = "No"
    If totalCases > 0 Then
        Dim avgDaily As Double: avgDaily = totalCases / 14
        Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
        Dim key As Variant, cnt As Long, maxDaily As Long, spikeDay As String
        maxDaily = 0
        For Each rw In tbl.ListRows
            If Not IsEmpty(rw.Range.Cells(1, tbl.ListColumns("TimeCreated").Index).Value) Then
                Dim dt As Date: dt = rw.Range.Cells(1, tbl.ListColumns("TimeCreated").Index).Value
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
        For Each key In dict.Keys
            cnt = CLng(dict(key))
            If cnt > maxDaily Then
                maxDaily = cnt
                spikeDay = CStr(key)
            End If
        Next key
        If maxDaily > 1.5 * avgDaily And maxDaily > 0 Then
            spikeMsg = "Yes (" & maxDaily & " on " & spikeDay & ")"
            LogEvent "Spike detected: " & maxDaily & " cases on " & spikeDay & "."
        End If
    End If
    Range("MetricSpike").Value = spikeMsg
    If spikeMsg <> "No" Then
        Range("MetricSpike").Font.Color = COLOR_HIGHLIGHT
    Else
        Range("MetricSpike").Font.Color = COLOR_TEXT
    End If
End Sub

Public Sub ToggleElement(elementName As String)
    Dim wsDash As Worksheet: Set wsDash = Worksheets("Dashboard")
    Dim shp As Shape
    On Error Resume Next
    Set shp = wsDash.Shapes(elementName)
    On Error GoTo 0
    If Not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        LogEvent "Toggled '" & elementName & "' to " & IIf(shp.Visible, "Shown", "Hidden") & "."
    Else
        MsgBox "Dashboard element '" & elementName & "' not found.", vbExclamation
    End If
End Sub

Private Sub LogEvent(message As String)
    Dim wsLog As Worksheet: Set wsLog = Worksheets("Log")
    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2
    wsLog.Cells(nextRow, 1).Value = Now
    wsLog.Cells(nextRow, 2).Value = message
    wsLog.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    wsLog.Columns(1).AutoFit
    wsLog.Columns(2).AutoFit
End Sub
