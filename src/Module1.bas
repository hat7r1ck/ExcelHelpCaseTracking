' Module: Module1.bas
' Description: Comprehensive module for tracking help cases and related metrics,
' including automatic sheet initialization, case logging with MTTP and MTTR,
' spike detection, inter-case gap calculation, and a Late Note Checker.
'
' Expected sheet structures:
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
'     Created automatically if missing

' ******************************
' Initialization Routine
' ******************************
Sub InitializeWorkbook()
    Dim ws As Worksheet
    
    ' Create Data_Import sheet if it doesn't exist
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Data_Import")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = "Data_Import"
    End If
    With ws
        .Range("A1").Value = "CaseID"
        .Range("B1").Value = "Owner"
        .Range("C1").Value = "TimeCreated"
        .Range("D1").Value = "OtherInfo"
        .Range("E1").Value = "TimeClosed"
    End With
    
    ' Create CaseLog sheet if it doesn't exist
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CaseLog")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "CaseLog"
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
    
    ' Create QuickEntry sheet if it doesn't exist
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("QuickEntry")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("CaseLog"))
        ws.Name = "QuickEntry"
    End If
    With ws
        .Range("A1").Value = "QuickEntry Sheet - Enter the CaseID in B2, Notes in B3, and your Owner ID in B4"
        .Range("B2").Value = ""
        .Range("B3").Value = ""
        .Range("B4").Value = ""
    End With
    
    ' Create Dashboard sheet if it doesn't exist
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("QuickEntry"))
        ws.Name = "Dashboard"
    End If
    With ws
        .Range("B1").Value = "Last Updated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    End With
    
    ' Create Log sheet if it doesn't exist
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Log")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Dashboard"))
        ws.Name = "Log"
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
' 1. Function: FindCaseInDataImport
' ******************************
Function FindCaseInDataImport(caseID As String) As Range
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("Data_Import")
    Set FindCaseInDataImport = wsData.Range("A:A").Find(What:=caseID, LookIn:=xlValues, LookAt:=xlWhole)
End Function

' ******************************
' 2. Update Data from CSV Export
' ******************************
Sub UpdateDataImportFromCSV(caseID As String)
    Dim csvPath As String
    Dim wbCSV As Workbook
    Dim wsCSV As Worksheet
    Dim rngFound As Range
    Dim wsData As Worksheet
    Dim lastRow As Long
    
    csvPath = "C:\Exports\XSOAR_Export.csv"
    
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
' 3. Operation Logging
' ******************************
Sub LogEvent(eventText As String)
    Dim wsLogSheet As Worksheet
    Dim nextRow As Long
    
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
' 4. Clear QuickEntry Fields
' ******************************
Sub ClearQuickEntry()
    Dim wsQuick As Worksheet
    Set wsQuick = ThisWorkbook.Worksheets("QuickEntry")
    wsQuick.Range("B2").Value = ""
    wsQuick.Range("B3").Value = ""
End Sub

' ******************************
' 5. Update Dashboard Timestamp
' ******************************
Sub UpdateDashboardTimestamp()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    wsDash.Range("B1").Value = "Last Updated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

' ******************************
' 6. Refresh All Data Connections
' ******************************
Sub RefreshAllData()
    ThisWorkbook.RefreshAll
    LogEvent "Data connections refreshed."
End Sub

' ******************************
' 7. Auto-Refresh Timer (Optional)
' ******************************
Public NextRefreshTime As Date

Sub StartAutoRefresh()
    NextRefreshTime = Now + TimeValue("00:05:00")
    Application.OnTime NextRefreshTime, "RefreshAllData"
    LogEvent "Auto-refresh scheduled at " & Format(NextRefreshTime, "hh:nn:ss")
End Sub

Sub StopAutoRefresh()
    On Error Resume Next
    Application.OnTime EarliestTime:=NextRefreshTime, Procedure:="RefreshAllData", Schedule:=False
    LogEvent "Auto-refresh stopped."
End Sub

' ******************************
' 8. Function: DetectSpike
' ******************************
Function DetectSpike(creationTime As Date) As Boolean
    Dim wsData As Worksheet
    Dim cell As Range
    Dim countSpike As Long
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
' 9. Function: GetLastClosedTime
' ******************************
Function GetLastClosedTime(ownerID As String, currentPickupTime As Date) As Variant
    Dim wsData As Worksheet
    Dim lastClosed As Date
    Dim cell As Range
    Dim currentDate As Date
    Dim found As Boolean
    
    Set wsData = ThisWorkbook.Worksheets("Data_Import")
    found = False
    lastClosed = 0
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, "B").End(xlUp).Row
    
    For Each cell In wsData.Range("B2:B" & lastRow)
        If LCase(cell.Value) = LCase(ownerID) Then
            If IsDate(cell.Offset(0, 3).Value) Then
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
' 10. Main Routine: AddHelpCase
' ******************************
Sub AddHelpCase()
    Dim wsData As Worksheet         ' Data_Import sheet
    Dim wsLogSheet As Worksheet     ' CaseLog sheet
    Dim wsQuick As Worksheet        ' QuickEntry sheet for input
    Dim caseID As String            ' Entered CaseID
    Dim noteText As String          ' Optional notes
    Dim ownerFromQuick As String    ' Owner from QuickEntry (B4)
    Dim foundCell As Range          ' Found case in Data_Import
    Dim nextRow As Long             ' Next available row in CaseLog
    
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
                MsgBox "Case " & caseID & " still not found after updating from CSV. Please re-run the macro later.", vbExclamation, "Case Not Found"
                LogEvent "Case " & caseID & " not found after CSV update. Manual re-check required."
                Exit Sub
            End If
        End If
    End If
    
    nextRow = wsLogSheet.Cells(wsLogSheet.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Log details into CaseLog:
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
            wsLogSheet.Cells(nextRow, "H").Interior.Color = vbYellow
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
    
    Dim spikeDetected As Boolean
    spikeDetected = DetectSpike(foundCell.Offset(0, 2).Value)
    If spikeDetected Then
        wsLogSheet.Cells(nextRow, "J").Value = "Spike Detected"
        wsLogSheet.Cells(nextRow, "J").Interior.Color = vbGreen
        LogEvent "Spike detected around case " & caseID
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
' 11. Late Note Checker Subroutine
' ******************************
Sub CheckLateNotes()
    Dim wsLog As Worksheet
    Dim lastRow As Long, i As Long
    Dim issuesCount As Long
    
    Set wsLog = ThisWorkbook.Worksheets("CaseLog")
    lastRow = wsLog.Cells(wsLog.Rows.Count, "H").End(xlUp).Row
    issuesCount = 0
    
    For i = 2 To lastRow
        If UCase(Trim(wsLog.Cells(i, "H").Value)) = "NOTE REQUIRED" Then
            issuesCount = issuesCount + 1
        End If
    Next i
    
    If issuesCount = 0 Then
        MsgBox "Congratulations, you rock!" & vbCrLf & "All cases that required notes have been addressed.", vbInformation, "Late Note Check"
    ElseIf issuesCount = 1 Then
        MsgBox "There is 1 case with a pending note. Please address it when possible.", vbExclamation, "Late Note Check"
    Else
        MsgBox "There are " & issuesCount & " cases with pending notes. Please address these issues.", vbExclamation, "Late Note Check"
    End If
End Sub
