Sub AddMissedCase()
    Dim wsLog As Worksheet
    Dim nextRow As Long
    Dim caseID As String, noteText As String, ownerID As String
    Dim pickupTime As Date, closedTime As Variant
    Dim response As Variant
    
    ' Set reference to your CaseLog sheet.
    Set wsLog = ThisWorkbook.Worksheets("CaseLog")
    
    ' Prompt for case details.
    caseID = InputBox("Enter CaseID for the missed case:", "Missed Case Entry")
    If caseID = "" Then Exit Sub
    
    ownerID = InputBox("Enter your Owner ID:", "Missed Case Entry")
    noteText = InputBox("Enter any notes (optional):", "Missed Case Entry")
    
    ' Use current time as default pickup time, allow override.
    pickupTime = Now
    response = InputBox("Enter pickup time (mm/dd/yyyy hh:mm) or leave blank for now:", "Missed Case Entry", Format(pickupTime, "mm/dd/yyyy hh:mm"))
    If Trim(response) <> "" And IsDate(response) Then
        pickupTime = CDate(response)
    End If
    
    ' For missed cases, assume closed time is not available.
    closedTime = "Open"
    
    ' Append a new row to CaseLog.
    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    wsLog.Cells(nextRow, "A").Value = caseID
    wsLog.Cells(nextRow, "B").Value = ownerID
    wsLog.Cells(nextRow, "C").Value = "N/A"          ' TimeCreated not available.
    wsLog.Cells(nextRow, "D").Value = pickupTime      ' QuickEntry Time.
    wsLog.Cells(nextRow, "E").Value = closedTime      ' TimeClosed.
    wsLog.Cells(nextRow, "F").Value = noteText          ' Notes.
    
    ' Set MTTP to "Backlogged" to mark it as a missed case.
    wsLog.Cells(nextRow, "G").Value = "Backlogged"
    
    ' Optionally, set Late Note Status to a default value.
    wsLog.Cells(nextRow, "H").Value = "Pending"
    
    MsgBox "Missed case added with MTTP marked as 'Backlogged'.", vbInformation, "Missed Case Entry"
End Sub
