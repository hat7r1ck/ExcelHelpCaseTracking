Sub UpdatePendingData()
    Dim wsData As Worksheet, wsLogSheet As Worksheet
    Dim lastRowCaseLog As Long, i As Long
    Dim caseID As String
    Dim foundCell As Range
    Dim updatedCount As Long
    Dim pickupDelay As Double
    
    Set wsData = ThisWorkbook.Worksheets("Data_Import")
    Set wsLogSheet = ThisWorkbook.Worksheets("CaseLog")
    lastRowCaseLog = wsLogSheet.Cells(wsLogSheet.Rows.Count, "A").End(xlUp).Row
    updatedCount = 0
    
    For i = 2 To lastRowCaseLog
        caseID = Trim(wsLogSheet.Cells(i, "A").Value)
        If caseID <> "" Then
            ' Check if either TimeCreated (Column C) or TimeClosed (Column E) is blank or has a placeholder ("DATA PENDING", "N/A", or "OPEN")
            If (Trim(UCase(wsLogSheet.Cells(i, "C").Value)) = "" Or _
                Trim(UCase(wsLogSheet.Cells(i, "C").Value)) = "DATA PENDING" Or _
                Trim(UCase(wsLogSheet.Cells(i, "C").Value)) = "N/A") Or _
               (Trim(UCase(wsLogSheet.Cells(i, "E").Value)) = "" Or _
                Trim(UCase(wsLogSheet.Cells(i, "E").Value)) = "DATA PENDING" Or _
                Trim(UCase(wsLogSheet.Cells(i, "E").Value)) = "OPEN") Then
                    
                Set foundCell = wsData.Range("A:A").Find(What:=caseID, LookIn:=xlValues, LookAt:=xlWhole)
                If Not foundCell Is Nothing Then
                    ' Update TimeCreated (Column C) from Data_Import (assumed at offset 2)
                    wsLogSheet.Cells(i, "C").Value = foundCell.Offset(0, 2).Value
                    
                    ' Update TimeClosed (Column E): if Data_Import has a valid date, update it; else set to "Open"
                    If IsDate(foundCell.Offset(0, 4).Value) Then
                        wsLogSheet.Cells(i, "E").Value = foundCell.Offset(0, 4).Value
                    Else
                        wsLogSheet.Cells(i, "E").Value = "Open"
                    End If
                    
                    ' Recalculate MTTP (Column G) if TimeCreated and QuickEntry Time (Column D) are dates
                    If IsDate(wsLogSheet.Cells(i, "C").Value) And IsDate(wsLogSheet.Cells(i, "D").Value) Then
                        wsLogSheet.Cells(i, "G").Value = FormatMinutes(DateDiff("n", wsLogSheet.Cells(i, "C").Value, wsLogSheet.Cells(i, "D").Value))
                    End If
                    
                    ' Determine pickup delay and update Late Note Status (Column H)
                    pickupDelay = DateDiff("n", wsLogSheet.Cells(i, "C").Value, wsLogSheet.Cells(i, "D").Value)
                    If pickupDelay >= 30 Then
                        If Len(Trim(wsLogSheet.Cells(i, "F").Value)) = 0 Then
                            wsLogSheet.Cells(i, "H").Value = "NOTE REQUIRED"
                        Else
                            wsLogSheet.Cells(i, "H").Value = "Note provided"
                        End If
                    Else
                        wsLogSheet.Cells(i, "H").Value = "On time"
                    End If
                    
                    ' Recalculate MTTR (Column I) if both TimeCreated and TimeClosed are valid dates
                    If IsDate(wsLogSheet.Cells(i, "C").Value) And IsDate(wsLogSheet.Cells(i, "E").Value) Then
                        wsLogSheet.Cells(i, "I").Value = FormatMinutes(DateDiff("n", wsLogSheet.Cells(i, "C").Value, wsLogSheet.Cells(i, "E").Value))
                    Else
                        wsLogSheet.Cells(i, "I").Value = "Open"
                    End If
                    
                    ' Update Spike Detection (Column J) based on TimeCreated
                    Dim spikeCount As Long
                    spikeCount = GetSpikeCount(wsLogSheet.Cells(i, "C").Value)
                    If spikeCount >= 5 Then
                        wsLogSheet.Cells(i, "J").Value = "Spike Detected (" & spikeCount & " cases)"
                    Else
                        wsLogSheet.Cells(i, "J").Value = "No spike"
                    End If
                    
                    ' Update Inter-case Gap (Column K) for owner cases using QuickEntry Time (Column D)
                    Dim ownerInLog As String
                    ownerInLog = Trim(wsLogSheet.Cells(i, "B").Value)
                    Dim lastClosedTime As Variant
                    lastClosedTime = GetLastClosedTime(ownerInLog, wsLogSheet.Cells(i, "D").Value)
                    If IsDate(lastClosedTime) Then
                        wsLogSheet.Cells(i, "K").Value = FormatMinutes(DateDiff("n", lastClosedTime, wsLogSheet.Cells(i, "D").Value))
                    Else
                        wsLogSheet.Cells(i, "K").Value = "N/A"
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
