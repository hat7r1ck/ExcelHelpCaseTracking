' Module: Module1.bas
' Description: Contains the VBA code to log a help case.
' This module retrieves a case ID from the QuickEntry sheet,
' looks up the corresponding case details in the Data_Import sheet,
' and logs the case (with a help timestamp) into the HelpCaseLog sheet.
'
' Optional: It also allows the user to add additional notes.

Sub AddHelpCase()
    ' Declare variables for worksheets and data.
    Dim wsData As Worksheet         ' Data source sheet (imported case details)
    Dim wsLog As Worksheet          ' Log sheet where help cases are recorded
    Dim wsQuick As Worksheet        ' QuickEntry sheet where user inputs are provided
    Dim caseID As String            ' The CaseID entered by the user
    Dim noteText As String          ' Optional notes entered by the user
    Dim foundCell As Range          ' The cell where the CaseID is found in the Data_Import sheet
    Dim nextRow As Long             ' Next available row in the HelpCaseLog sheet

    ' Set worksheet variables – update sheet names if needed.
    Set wsData = ThisWorkbook.Worksheets("Data_Import")
    Set wsLog = ThisWorkbook.Worksheets("HelpCaseLog")
    Set wsQuick = ThisWorkbook.Worksheets("QuickEntry")
    
    ' Get the CaseID from the QuickEntry sheet (assumes cell B2 contains the input)
    caseID = Trim(wsQuick.Range("B2").Value)
    
    ' Get optional notes from cell B3 in the QuickEntry sheet (if you want to log extra info)
    noteText = Trim(wsQuick.Range("B3").Value)
    
    ' Validate that a CaseID was provided.
    If caseID = "" Then
        MsgBox "Please enter a CaseID in cell B2.", vbExclamation, "Missing Input"
        Exit Sub
    End If
    
    ' Attempt to locate the CaseID in the Data_Import sheet.
    ' Assumes that the CaseID is located in column A.
    Set foundCell = wsData.Range("A:A").Find(What:=caseID, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        ' Determine the next empty row in the HelpCaseLog sheet.
        ' This assumes that row 1 is a header row.
        nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
        
        ' Log the case details into the HelpCaseLog sheet.
        ' Column mapping in HelpCaseLog:
        '   Column A: CaseID
        '   Column B: TimeCreated (from Data_Import, assumed to be in column B)
        '   Column C: HelpTimestamp (current time)
        '   Column D: TimeClosed (from Data_Import, assumed to be in column D)
        '   Column E: Notes (optional, from QuickEntry)
        wsLog.Cells(nextRow, "A").Value = foundCell.Value                        ' CaseID
        wsLog.Cells(nextRow, "B").Value = foundCell.Offset(0, 1).Value             ' TimeCreated from Data_Import (Column B)
        wsLog.Cells(nextRow, "C").Value = Now                                      ' HelpTimestamp – logs the current time
        wsLog.Cells(nextRow, "D").Value = foundCell.Offset(0, 3).Value             ' TimeClosed from Data_Import (Column D)
        wsLog.Cells(nextRow, "E").Value = noteText                                 ' Optional notes from QuickEntry
        
        ' Notify the user that the case was logged successfully.
        MsgBox "Case " & caseID & " logged successfully with a help timestamp.", vbInformation, "Success"
    Else
        ' If the case could not be found, notify the user.
        MsgBox "CaseID " & caseID & " was not found in the Data_Import sheet.", vbExclamation, "Case Not Found"
    End If
End Sub
