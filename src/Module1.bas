' Module: Module1.bas
' Description: Comprehensive module for tracking help cases and related metrics,
' including automatic sheet initialization.
'
' Required sheets and headers:
'   Data_Import:
'     A: CaseID, B: Owner, C: TimeCreated, D: OtherInfo, E: TimeClosed
'
'   CaseLog:
'     A: CaseID, B: Owner, C: TimeCreated, D: QuickEntry Time, E: TimeClosed,
'     F: Notes, G: MTTP, H: Late Note Status, I: MTTR, J: Spike Detection, K: Inter-case Gap
'
'   QuickEntry:
'     B2: CaseID input, B3: Notes input, B4: Owner (entered by user)
'
'   Dashboard:
'     B1: "Last Updated" timestamp
'
'   Log:
'     A: Timestamp, B: Event (created automatically if missing)

Option Explicit

Public NextRefreshTime As Date

' ******************************
' Helper Function: SheetExists
' ******************************
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' ******************************
' Initialization Routine
' ******************************
Sub InitializeWorkbook()
    Dim ws As Worksheet
    
    ' Data_Import
    If Not SheetExists("Data_Import") Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = "Data_Import"
    Else
        Set ws = ThisWorkbook.Worksheets("Data_Import")
    End If
    With ws
        .Range("A1").Value = "CaseID"
        .Range("B1").Value = "Owner"
        .Range("C1").Value = "TimeCreated"
        .Range("D1").Value = "OtherInfo"
        .Range("E1").Value = "TimeClosed"
    End With
    
    ' CaseLog
    If Not SheetExists("CaseLog") Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "CaseLog"
    Else
        Set ws = ThisWorkbook.Worksheets("CaseLog")
    End If
    With ws
        .Range("A1").Value = "CaseID"
        .Range("B1").Value = "Owner"
        .Range("C1").Value = "TimeCreated"
        .Range("D1").Value = "QuickEntry Time"
        .Range("E1").Value = "TimeClosed"
        .Range("F1").Value = "Notes"
        .Range("G1").Value = "MTTP"
        .Range("H1").Value = "Late Note Status"
        .Range("I1").Value = "MTTR"
        .Range("J1").Value = "Spike Detection"
        .Range("K1").Value = "Inter-case Gap"
    End With
    
    ' QuickEntry
    If Not SheetExists("QuickEntry") Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("CaseLog"))
        ws.Name = "QuickEntry"
    Else
        Set ws = ThisWorkbook.Worksheets("QuickEntry")
    End If
    With ws
        .Range("A1").Value = "QuickEntry Sheet - Enter CaseID in B2, Notes in B3, and Owner ID in B4"
        .Range("B2").Value = ""
        .Range("B3").Value = ""
        .Range("B4").Value = ""
    End With
    
    ' Dashboard
    If Not SheetExists("Dashboard") Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("QuickEntry"))
        ws.Name = "Dashboard"
    Else
        Set ws = ThisWorkbook.Worksheets("Dashboard")
    End If
    With ws
        .Range("B1").Value = "Last Updated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    End With
    
    ' Log
    If Not SheetExists("Log") Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Dashboard"))
        ws.Name = "Log"
    Else
        Set ws = ThisWorkbook.Worksheets("Log")
    End If
    With ws
        .Range("A1").Value = "Timestamp"
        .Range("B1").Value = "Event"
    End With
    
    MsgBox "Workbook initialized with required sheets and headers.", vbInformation, "Initialization Complete"
End Sub

' ******************************
' Helper Function: FormatMinutes
' ******************************
Function FormatMinutes(mins As Double) As String
    Dim hrs As Long, remMins As Long
    If mins < 0 Then
        FormatMinutes = "-" & FormatMinutes(-mins)
    ElseIf mins < 60 Then
        FormatMinutes = mins & " mins"
    Else
        hrs = mins \ 60
        remMins = mins Mod 60
        FormatMinutes = hrs & " hrs " & remMins & " mins"
    End If
End Function

' ******************************
' Function: FindCaseInDataImport
' ******************************
Function FindCaseInDataImport(caseID As String) As Range
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("Data_Import")
    Set FindCaseInDataImport = wsData.Range("A:A").Find(What:=caseID, LookIn:=xlValues, LookAt:=xlWhole)
End Function

' ******************************
' Sub: UpdateDataImportFromCSV
' ******************************
Sub UpdateDataImportFromCSV(caseID As String)
    Dim csvPath As String
    Dim wbCSV As Workbook, wsCSV As Worksheet, rngFound As Range
    Dim wsData As Worksheet, lastRow As Long
    
    csvPath = "C:\Exports\XSOAR_Export.csv"  ' Adjust as needed.
    
    On Error Resume Next
    Set wbCSV = Workbooks.Open(csvPath)
    On Error GoTo 0
    If wbCSV Is Nothing Then
        MsgBox "CSV file not found at " & csvPath, vbExclamation
        LogEvent "CSV file not found at " & csvPath
        Exit Sub
    End If
    
    Set wsCSV = wbCSV.Worksheets(1)
    Set rngFound = wsCSV.Range("A:A").Find(What:=caseID, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngFound Is Nothing Then
        Set wsData = ThisWorkbook.Worksheets("Data_Import")
        lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row + 1
        wsCSV.Rows(rngFound.Row).Copy wsData.Rows(lastRow)
        LogEvent "Updated Data_Import with case " & caseID & " from CSV export."
    Else
        MsgBox "Case ID " & caseID & " not found in CSV export.", vbExclamation
        LogEvent "Case " & caseID & " not found in CSV export."
    End If
    wbCSV.Close SaveChanges:=False
End Sub

' ******************************
' Sub: LogEvent
' ******************************
Sub LogEvent(eventText As String)
    Dim wsLogSheet As Worksheet, nextRow As Long
    On Error Resume Next
    Set wsLogSheet = ThisWorkbook.Worksheets("Log")
    On Error GoTo 0
    If wsLogSheet Is Nothing Then
        Set wsLogSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsLogSheet.Name = "Log"
        wsLogSheet.Range("A1").Value = "Timestamp"
        wsLogSheet.Range("B1").Value = "Event"
    End If
    nextRow = wsLogSheet.Cells(wsLogSheet.Rows.Count, "A").End(xlUp).Row + 1
    wsLogSheet.Cells(nextRow, "A").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    wsLogSheet.Cells(nextRow, "B").Value = eventText
End Sub

' ******************************
' Sub: ClearQuickEntry
' ******************************
Sub ClearQuickEntry()
    Dim wsQuick As Worksheet
    Set wsQuick = ThisWorkbook.Worksheets("QuickEntry")
    wsQuick.Range("B2").Value = ""
    wsQuick.Range("B3").Value = ""
End Sub

' ******************************
' Sub: UpdateDashboardTimestamp
' ******************************
Sub UpdateDashboardTimestamp()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    wsDash.Range("B1").Value = "Last Updated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

' ******************************
' Sub: RefreshAllData
' ******************************
Sub RefreshAllData()
    ThisWorkbook.RefreshAll
    LogEvent "Data connections refreshed."
End Sub

' ******************************
' Sub: StartAutoRefresh
' ******************************
Sub StartAutoRefresh()
    NextRefreshTime = Now + TimeValue("00:05:00")
    Application.OnTime NextRefreshTime, "RefreshAllData"
    LogEvent "Auto-refresh scheduled at " & Format(NextRefreshTime, "hh:nn:ss")
End Sub

' ******************************
' Sub: StopAutoRefresh
' ******************************
Sub StopAutoRefresh()
    On Error Resume Next
    Application.OnTime EarliestTime:=NextRefreshTime, Procedure:="RefreshAllData", Schedule:=False
    LogEvent "Auto-refresh stopped."
End Sub

' ******************************
' Function: DetectSpike
' ******************************
Function DetectSpike(creationTime As Date) As Boolean
    Dim wsData As Worksheet, cell As Range, countSpike As Long
    Dim lowerBound As Date, upperBound As Date
    lowerBound = DateAdd("n", -5, creationTime)
    upperBound = DateAdd("n", 5, creationTime)
    Set wsData = ThisWorkbook.Worksheets("Data_Import")
    countSpike = 0
    For Each cell In wsData.Range("C2:C" & wsData.Cells(wsData.Rows.Count, "C").End(xlUp).Row)
        If IsDate(cell.Value) Then
            If cell.Value >= lowerBound And cell.Value <= upperBound Then
                countSpike = countSpike + 1
            End If
        End If
    Next cell
    DetectSpike = (countSpike >= 5)
End Function

' ******************************
' Function: GetSpikeCount
' ******************************
Function GetSpikeCount(creationTime As Date) As Long
    Dim wsData As Worksheet, cell As Range, countSpike As Long
    Dim lowerBound As Date, upperBound As Date
    lowerBound = DateAdd("n", -5, creationTime)
    upperBound = DateAdd("n", 5, creationTime)
    Set wsData = ThisWorkbook.Worksheets("Data_Import")
    countSpike = 0
    For Each cell In wsData.Range("C2:C" & wsData.Cells(wsData.Rows.Count, "C").End(xlUp).Row)
        If IsDate(cell.Value) Then
            If cell.Value >= lowerBound And cell.Value <= upperBound Then
                countSpike = countSpike + 1
            End If
        End If
    Next cell
    GetSpikeCount = countSpike
End Function

' ******************************
' Function: GetLastClosedTime
' ******************************
Function GetLastClosedTime(ownerID As String, currentPickupTime As Date) As Variant
    Dim wsData As Worksheet, lastClosed As Date, cell As Range
    Dim currentDate As Date, found As Boolean
    Set wsData = ThisWorkbook.Worksheets("Data_Import")
    found = False
    lastClosed = 0
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, "B").End(xlUp).Row
    For Each cell In wsData.Range("B2:B" & lastRow)
        If LCase(cell.Value) = LCase(ownerID) Then
            If IsDate(cell.Offset(0, 3).Value) Then  ' TimeClosed in column E
                currentDate = cell.Offset(0, 3).Value
                If currentDate < currentPickupTime Then
                    If Not found Then
                        lastClosed = currentDate
                        found = True
                    ElseIf currentDate > lastClosed Then
                        lastClosed = currentDate
                    End If
                End If
            End If
        End If
    Next cell
    If found Then
        GetLastClosedTime = lastClosed
    Else
        GetLastClosedTime = CVErr(xlErrNA)
    End If
End Function

' ******************************
' Main Routine: AddHelpCase
' ******************************
Sub AddHelpCase()
    Dim wsData As Worksheet, wsLogSheet As Worksheet, wsQuick As Worksheet
    Dim caseID As String, noteText As String, ownerFromQuick As String
    Dim foundCell As Range, nextRow As Long
    
    Set wsData = ThisWorkbook.Worksheets("Data_Import")
    Set wsLogSheet = ThisWorkbook.Worksheets("CaseLog")
    Set wsQuick = ThisWorkbook.Worksheets("QuickEntry")
    
    caseID = Trim(wsQuick.Range("B2").Value)
    noteText = Trim(wsQuick.Range("B3").Value)
    ownerFromQuick = Trim(wsQuick.Range("B4").Value)
    
    If caseID = "" Then
        MsgBox "Please enter a CaseID in cell B2.", vbExclamation, "Missing Input"
        Exit Sub
    End If
    If ownerFromQuick = "" Then
        MsgBox "Please enter your Owner ID in cell B4.", vbExclamation, "Missing Owner"
        Exit Sub
    End If
    
    LogEvent "Processing CaseID: " & caseID
    
    Set foundCell = FindCaseInDataImport(caseID)
    If foundCell Is Nothing Then
        RefreshAllData
        Application.Wait Now + TimeValue("00:00:05")
        Set foundCell = FindCaseInDataImport(caseID)
        If foundCell Is Nothing Then
            UpdateDataImportFromCSV caseID
            Set foundCell = FindCaseInDataImport(caseID)
            If foundCell Is Nothing Then
                ' CASE NOT FOUND: Add a placeholder row to the CaseLog
                nextRow = wsLogSheet.Cells(wsLogSheet.Rows.Count, "A").End(xlUp).Row + 1
                wsLogSheet.Cells(nextRow, "A").Value = caseID
                wsLogSheet.Cells(nextRow, "B").Value = ownerFromQuick
                wsLogSheet.Cells(nextRow, "C").Value = "N/A"          ' TimeCreated not available
                wsLogSheet.Cells(nextRow, "D").Value = Now             ' QuickEntry Time
                wsLogSheet.Cells(nextRow, "E").Value = "Data pending"  ' TimeClosed not available
                wsLogSheet.Cells(nextRow, "F").Value = noteText
                wsLogSheet.Cells(nextRow, "G").Value = "Data pending"  ' MTTP not available
                wsLogSheet.Cells(nextRow, "H").Value = "Data pending"  ' Late Note Status not available
                wsLogSheet.Cells(nextRow, "I").Value = "Data pending"  ' MTTR not available
                wsLogSheet.Cells(nextRow, "J").Value = "Data pending"  ' Spike Detection not available
                wsLogSheet.Cells(nextRow, "K").Value = "Data pending"  ' Inter-case Gap not available
                MsgBox "Case " & caseID & " not found in Data_Import. Placeholder row added. Please update data later.", vbExclamation, "Case Not Found"
                LogEvent "Case " & caseID & " placeholder added due to missing data."
                ClearQuickEntry
                UpdateDashboardTimestamp
                Exit Sub
            End If
        End If
    End If
    
    nextRow = wsLogSheet.Cells(wsLogSheet.Rows.Count, "A").End(xlUp).Row + 1
    
    wsLogSheet.Cells(nextRow, "A").Value = foundCell.Value                      ' CaseID
    wsLogSheet.Cells(nextRow, "B").Value = foundCell.Offset(0, 1).Value           ' Owner
    wsLogSheet.Cells(nextRow, "C").Value = foundCell.Offset(0, 2).Value           ' TimeCreated
    wsLogSheet.Cells(nextRow, "D").Value = Now                                    ' QuickEntry Time
    If IsDate(foundCell.Offset(0, 4).Value) Then
        wsLogSheet.Cells(nextRow, "E").Value = foundCell.Offset(0, 4).Value       ' TimeClosed
    Else
        wsLogSheet.Cells(nextRow, "E").Value = "Open"
    End If
    wsLogSheet.Cells(nextRow, "F").Value = noteText                               ' Notes
    
    Dim pickupDelay As Double
    pickupDelay = DateDiff("n", foundCell.Offset(0, 2).Value, Now)
    wsLogSheet.Cells(nextRow, "G").Value = FormatMinutes(pickupDelay)
    
    If pickupDelay >= 30 Then
        If Len(noteText) = 0 Then
            MsgBox "This case was picked up " & pickupDelay & " minutes after creation. Please add a note for discussion.", vbExclamation, "Late Pickup"
            wsLogSheet.Cells(nextRow, "H").Value = "NOTE REQUIRED"
            LogEvent "Case " & caseID & " picked up after " & pickupDelay & " minutes; note required."
        Else
            wsLogSheet.Cells(nextRow, "H").Value = "Note provided"
        End If
    Else
        wsLogSheet.Cells(nextRow, "H").Value = "On time"
    End If
    
    Dim resolutionTime As Variant
    If IsDate(foundCell.Offset(0, 2).Value) And IsDate(foundCell.Offset(0, 4).Value) Then
        resolutionTime = DateDiff("n", foundCell.Offset(0, 2).Value, foundCell.Offset(0, 4).Value)
        wsLogSheet.Cells(nextRow, "I").Value = FormatMinutes(CDbl(resolutionTime))
    Else
        wsLogSheet.Cells(nextRow, "I").Value = "Open"
    End If
    
    Dim spikeCount As Long
    spikeCount = GetSpikeCount(foundCell.Offset(0, 2).Value)
    If spikeCount >= 5 Then
        wsLogSheet.Cells(nextRow, "J").Value = "Spike Detected (" & spikeCount & " cases)"
        LogEvent "Spike detected around case " & caseID & " (" & spikeCount & " cases)"
    Else
        wsLogSheet.Cells(nextRow, "J").Value = "No spike"
    End If
    
    Dim lastClosedTime As Variant
    If LCase(foundCell.Offset(0, 1).Value) = LCase(ownerFromQuick) Then
        lastClosedTime = GetLastClosedTime(ownerFromQuick, Now)
        If IsDate(lastClosedTime) Then
            wsLogSheet.Cells(nextRow, "K").Value = FormatMinutes(DateDiff("n", lastClosedTime, Now))
        Else
            wsLogSheet.Cells(nextRow, "K").Value = "N/A"
        End If
    Else
        wsLogSheet.Cells(nextRow, "K").Value = "N/A"
    End If
    
    MsgBox "Case " & caseID & " logged successfully with a QuickEntry Time.", vbInformation, "Success"
    LogEvent "Case " & caseID & " logged successfully."
    
    ClearQuickEntry
    UpdateDashboardTimestamp
End Sub

' ******************************
' Late Note Checker Subroutine
' ******************************
Sub CheckLateNotes()
    Dim wsLog As Worksheet
    Dim lastRow As Long, i As Long
    Dim issuesCount As Long
    
    Set wsLog = ThisWorkbook.Worksheets("CaseLog")
    lastRow = wsLog.Cells(wsLog.Rows.Count, "H").End(xlUp).Row
    issuesCount = 0
    
    For i = 2 To lastRow
        Dim status As String
        status = UCase(Trim(wsLog.Cells(i, "H").Value))
        If status = "NOTE REQUIRED" Or status = "DATA PENDING" Then
            issuesCount = issuesCount + 1
        End If
    Next i
    
    If issuesCount = 0 Then
        MsgBox "Congratulations, you rock!" & vbCrLf & "All cases that required notes have been addressed.", vbInformation, "Late Note Check"
    ElseIf issuesCount = 1 Then
        MsgBox "There is 1 case with a pending note or data issue. Please address it when possible.", vbExclamation, "Late Note Check"
    Else
        MsgBox "There are " & issuesCount & " cases with pending notes or data issues. Please address these issues.", vbExclamation, "Late Note Check"
    End If
End Sub

' ******************************
' Update Pending Data
' ******************************
Sub UpdatePendingData()
    Dim wsData As Worksheet, wsCaseLog As Worksheet
    Dim lastRowCaseLog As Long, i As Long
    Dim caseID As String
    Dim foundCell As Range
    Dim updatedCount As Long
    Dim pickupDelay As Double
    
    Set wsData = ThisWorkbook.Worksheets("Data_Import")
    Set wsCaseLog = ThisWorkbook.Worksheets("CaseLog")
    lastRowCaseLog = wsCaseLog.Cells(wsCaseLog.Rows.Count, "A").End(xlUp).Row
    updatedCount = 0
    
    ' Loop through each row in CaseLog (starting at row 2, assuming headers in row 1)
    For i = 2 To lastRowCaseLog
        caseID = Trim(wsCaseLog.Cells(i, "A").Value)
        If caseID <> "" Then
            ' Check if either TimeCreated (Column C) or TimeClosed (Column E) is blank or a placeholder
            If (Trim(UCase(wsCaseLog.Cells(i, "C").Value)) = "" Or _
                Trim(UCase(wsCaseLog.Cells(i, "C").Value)) = "DATA PENDING" Or _
                Trim(UCase(wsCaseLog.Cells(i, "C").Value)) = "N/A") Or _
               (Trim(UCase(wsCaseLog.Cells(i, "E").Value)) = "" Or _
                Trim(UCase(wsCaseLog.Cells(i, "E").Value)) = "DATA PENDING") Then
                    
                Set foundCell = wsData.Range("A:A").Find(What:=caseID, LookIn:=xlValues, LookAt:=xlWhole)
                If Not foundCell Is Nothing Then
                    ' Update TimeCreated (Column C) from Data_Import's Column C
                    wsCaseLog.Cells(i, "C").Value = foundCell.Offset(0, 2).Value
                    ' Update TimeClosed (Column E) from Data_Import's Column E if available; else "Open"
                    If IsDate(foundCell.Offset(0, 4).Value) Then
                        wsCaseLog.Cells(i, "E").Value = foundCell.Offset(0, 4).Value
                    Else
                        wsCaseLog.Cells(i, "E").Value = "Open"
                    End If
                    
                    ' Recalculate MTTP (Column G) if TimeCreated and QuickEntry Time (Column D) are valid dates
                    If IsDate(wsCaseLog.Cells(i, "C").Value) And IsDate(wsCaseLog.Cells(i, "D").Value) Then
                        wsCaseLog.Cells(i, "G").Value = FormatMinutes(DateDiff("n", wsCaseLog.Cells(i, "C").Value, wsCaseLog.Cells(i, "D").Value))
                    End If
                    
                    ' Recalculate Late Note Status (Column H) based on MTTP
                    pickupDelay = DateDiff("n", wsCaseLog.Cells(i, "C").Value, wsCaseLog.Cells(i, "D").Value)
                    If pickupDelay >= 30 Then
                        If Len(Trim(wsCaseLog.Cells(i, "F").Value)) = 0 Then
                            wsCaseLog.Cells(i, "H").Value = "NOTE REQUIRED"
                        Else
                            wsCaseLog.Cells(i, "H").Value = "Note provided"
                        End If
                    Else
                        wsCaseLog.Cells(i, "H").Value = "On time"
                    End If
                    
                    ' Recalculate MTTR (Column I) if both TimeCreated and TimeClosed are valid dates
                    If IsDate(wsCaseLog.Cells(i, "C").Value) And IsDate(wsCaseLog.Cells(i, "E").Value) Then
                        wsCaseLog.Cells(i, "I").Value = FormatMinutes(DateDiff("n", wsCaseLog.Cells(i, "C").Value, wsCaseLog.Cells(i, "E").Value))
                    End If
                    
                    ' Recalculate Spike Detection (Column J) based on updated TimeCreated
                    Dim spikeCount As Long
                    spikeCount = GetSpikeCount(wsCaseLog.Cells(i, "C").Value)
                    If spikeCount >= 5 Then
                        wsCaseLog.Cells(i, "J").Value = "Spike Detected (" & spikeCount & " cases)"
                    Else
                        wsCaseLog.Cells(i, "J").Value = "No spike"
                    End If
                    
                    ' Recalculate Inter-case Gap (Column K) for owner cases using QuickEntry Time (Column D)
                    Dim ownerInLog As String
                    ownerInLog = Trim(wsCaseLog.Cells(i, "B").Value)
                    Dim lastClosedTime As Variant
                    lastClosedTime = GetLastClosedTime(ownerInLog, wsCaseLog.Cells(i, "D").Value)
                    If IsDate(lastClosedTime) Then
                        wsCaseLog.Cells(i, "K").Value = FormatMinutes(DateDiff("n", lastClosedTime, wsCaseLog.Cells(i, "D").Value))
                    Else
                        wsCaseLog.Cells(i, "K").Value = "N/A"
                    End If
                    
                    updatedCount = updatedCount + 1
                End If
            End If
        End If
    Next i
    
    If updatedCount = 0 Then
        MsgBox "Update complete. No pending rows were updated.", vbInformation, "Update Pending Data"
    ElseIf updatedCount = 1 Then
        MsgBox "Update complete. 1 pending row was updated with new data.", vbInformation, "Update Pending Data"
    Else
        MsgBox "Update complete. " & updatedCount & " pending rows were updated with new data.", vbInformation, "Update Pending Data"
    End If
End Sub