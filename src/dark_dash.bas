Option Explicit

Public Sub CreateDarkModeDashboard()
    Dim dashSheet As Worksheet, dataSheet As Worksheet, logSheet As Worksheet
    Dim pvtCache As PivotCache, pvt As PivotTable
    Dim chartObj As ChartObject
    Dim slcCache As SlicerCache, slcObj As Slicer
    Dim tlCache As SlicerCache, tlObj As Slicer
    Dim lastRow As Long, lastCol As Long
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    ' **1. Create/Reference required sheets** 
    ' Data Sheet
    Set dataSheet = Nothing
    On Error Resume Next
    Set dataSheet = ThisWorkbook.Worksheets("Data")
    On Error GoTo ErrorHandler
    If dataSheet Is Nothing Then
        Set dataSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        dataSheet.Name = "Data"
    End If
    ' Dashboard Sheet
    Set dashSheet = Nothing
    On Error Resume Next
    Set dashSheet = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo ErrorHandler
    If dashSheet Is Nothing Then
        Set dashSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        dashSheet.Name = "Dashboard"
    End If
    ' Apply dark theme to entire Dashboard sheet
    dashSheet.Cells.Font.Color = RGB(240, 240, 240)     ' light font
    dashSheet.Cells.Interior.Color = RGB(45, 45, 48)    ' dark background
    dashSheet.Tab.Color = RGB(60, 60, 60)               ;' optional: dark tab color
    
    ' Log Sheet
    Set logSheet = Nothing
    On Error Resume Next
    Set logSheet = ThisWorkbook.Worksheets("Log")
    On Error GoTo ErrorHandler
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        logSheet.Name = "Log"
        ' Initialize log headers
        logSheet.Range("A1:B1").Value = Array("Timestamp", "Event")
    End If
    
    ' **2. Refresh data connections (ensure Data sheet is up-to-date)** 
    ThisWorkbook.RefreshAll
    ' (Optionally, wait for refresh to complete if asynchronous)
    Application.Wait (Now + TimeValue("0:00:01"))  ' small pause to allow refresh
    LogEvent logSheet, "Data connections refreshed"
    
    ' If data is not in a table, define the used range or convert to table
    With dataSheet
        If .UsedRange.Count > 1 Then  ' if there's data
            lastRow = .Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            lastCol = .Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        Else
            lastRow = 1
            lastCol = 1
        End If
        Dim dataRange As Range
        Set dataRange = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))
        ' Convert to Table (ListObject) if not already a table
        Dim dataTbl As ListObject
        On Error Resume Next
        Set dataTbl = .ListObjects("DataTable")
        On Error GoTo ErrorHandler
        If dataTbl Is Nothing Then
            If lastRow > 1 Or lastCol > 1 Then  ' if actual data exists beyond a single cell
                Set dataTbl = .ListObjects.Add(xlSrcRange, dataRange, , xlYes)
            Else
                ' Create an empty table with headers if only one cell (no data yet)
                .Cells(1, 1).Value = "DataHeader"
                Set dataTbl = .ListObjects.Add(xlSrcRange, .Range("A1:A2"), , xlYes)
                dataTbl.DataBodyRange.Delete  ' remove the dummy row
            End If
            dataTbl.Name = "DataTable"
        End If
        ' Ensure the named range "DataRange" refers to the full table data for backward compatibility
        On Error Resume Next
        ThisWorkbook.Names.Add Name:="DataRange", RefersTo:=dataTbl.Range
        On Error GoTo ErrorHandler
    End With
    
    ' **3. Create or Update Pivot Table on Dashboard** 
    On Error Resume Next
    Set pvt = dashSheet.PivotTables("DashboardPivot")
    On Error GoTo ErrorHandler
    If pvt Is Nothing Then
        ' Create new PivotCache from the data table 
        Set pvtCache = ThisWorkbook.PivotCaches.Create(xlDatabase, "DataTable", xlPivotTableVersion15)
        ' Create PivotTable on Dashboard sheet starting at cell A5
        Set pvt = pvtCache.CreatePivotTable(TableDestination:=dashSheet.Range("A5"), TableName:="DashboardPivot")
        ' Add fields to PivotTable
        pvt.ManualUpdate = True  ' defer updates for performance
        On Error Resume Next
        pvt.PivotFields("Category").Orientation = xlRowField
        pvt.PivotFields("Category").Position = 1
        pvt.PivotFields("Value").Orientation = xlDataField
        pvt.PivotFields("Value").Function = xlSum
        On Error GoTo 0
        pvt.ManualUpdate = False  ' now apply the changes
    Else
        ' Pivot exists – just refresh its cache to update data
        Set pvtCache = pvt.PivotCache
        pvtCache.Refresh
    End If
    ' PivotTable formatting and options
    If Not pvt Is Nothing Then
        pvt.PivotCache.EnableRefresh = True
        pvt.PreserveFormatting = True        ' preserve dark formatting on refresh [oai_citation_attribution:2‡learn.microsoft.com](https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable.preserveformatting#:~:text=True%20if%20formatting%20is%20preserved,or%20changing%20page%20field%20items)
        pvt.TableStyle2 = "PivotStyleDark1"   ' apply a dark pivot table style
        pvt.ColumnGrand = True: pvt.RowGrand = True  ' show totals (if needed)
        pvt.HasAutoFormat = False            ' prevent auto-formatting override
        pvt.DisplayErrorString = True: pvt.ErrorString = "–"  ' handle errors in data display
    End If
    
    ' **4. Create or Update Pivot Chart** 
    On Error Resume Next
    Set chartObj = dashSheet.ChartObjects("SalesChart")
    On Error GoTo ErrorHandler
    If chartObj Is Nothing Then
        Set chartObj = dashSheet.ChartObjects.Add(Left:=300, Top:=50, Width:=500, Height:=300)
        chartObj.Name = "SalesChart"
    End If
    With chartObj.Chart
        .SetSourceData Source:=pvt.TableRange2   ' link chart to entire pivot table range
        .ChartType = xlColumnClustered
        ' Apply dark theme to chart elements
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(45, 45, 48)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(45, 45, 48)
        .ChartArea.Format.Line.Visible = msoFalse
        ' Format chart title
        .HasTitle = True
        .ChartTitle.Text = "Sales by Category"
        .ChartTitle.Font.Color = RGB(240, 240, 240)
        ' Format axes and legend
        If .HasLegend Then .Legend.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(240, 240, 240)
        On Error Resume Next
        .Axes(xlCategory).TickLabels.Font.Color = RGB(240, 240, 240)
        .Axes(xlValue).TickLabels.Font.Color = RGB(240, 240, 240)
        On Error GoTo 0
        .Refresh   ' ensure the chart displays the latest pivot data
    End With
    
    ' **5. Create or Update Slicer and Timeline** 
    ' Slicer for Category field
    On Error Resume Next
    Set slcCache = ThisWorkbook.SlicerCaches("Slicer_Category")
    On Error GoTo ErrorHandler
    If slcCache Is Nothing Then
        ' Add slicer cache for the pivot field "Category"
        Set slcCache = ThisWorkbook.SlicerCaches.Add(pvt, "Category", "Slicer_Category", xlSlicer)
        ' Add the slicer to the dashboard sheet
        Set slcObj = slcCache.Slicers.Add(dashSheet, , "Slicer_Category", "Category", dashSheet.Range("H5"))
        slcObj.Style = "SlicerStyleDark1"  ' apply dark style to slicer
    Else
        ' If slicer exists, ensure it still points to the current pivot (it should if cache reused)
        Set slcObj = slcCache.Slicers(1)
    End If
    ' Timeline slicer for Date field (if exists in pivot)
    On Error Resume Next
    If Not pvt.PivotFields("Date") Is Nothing Then
        Set tlCache = ThisWorkbook.SlicerCaches("Timeline_Date")
        On Error GoTo ErrorHandler
        If tlCache Is Nothing Then
            Set tlCache = ThisWorkbook.SlicerCaches.Add(pvt, "Date", "Timeline_Date", xlTimeline)
            Set tlObj = tlCache.Slicers.Add(dashSheet, , "Timeline_Date", "Date", dashSheet.Range("H15"))
            tlObj.Style = "TimeSlicerStyleDark1"  ' apply dark style to timeline
        End If
    End If
    On Error GoTo ErrorHandler  ' reset error handling
    
    ' **6. Log completion event** 
    LogEvent logSheet, "Dashboard updated successfully"
    
    ' **7. Restore Excel settings and exit** 
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    ' Handle errors: log and alert the user
    Dim errMsg As String
    errMsg = "Error " & Err.Number & ": " & Err.Description
    If Not logSheet Is Nothing Then
        LogEvent logSheet, errMsg
    End If
    MsgBox errMsg, vbCritical, "Dashboard Module Error"
    ' Restore settings even if there is an error
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' Helper subroutine for logging events to the Log sheet
Private Sub LogEvent(ws As Worksheet, msg As String)
    On Error Resume Next
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(lr, "A").Value = Now    ' Timestamp
    ws.Cells(lr, "B").Value = msg   ' Event message
End Sub
