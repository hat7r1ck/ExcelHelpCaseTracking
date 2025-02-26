Sub UpdatePendingData()
    Dim wsData As Worksheet, wsLogSheet As Worksheet
    Dim lastRowCaseLog As Long, i As Long
    Dim caseID As String
    Dim foundCell As Range
    Dim updatedCount As Long
    Dim pickupDelay As Double
    Dim rowUpdated As Boolean

    Set wsData = ThisWorkbook.Worksheets("Data_Import")
    Set wsLogSheet = ThisWorkbook.Worksheets("CaseLog")
    lastRowCaseLog = wsLogSheet.Cells(wsLogSheet.Rows.Count, "A").End(xlUp).Row
    updatedCount = 0

    For i = 2 To lastRowCaseLog
        rowUpdated = False
        caseID = Trim(wsLogSheet.Cells(i, "A").Value)
        If caseID <> "" Then
            ' Check if either TimeCreated (col C) or TimeClosed (col E) is blank or a placeholder
            If (Trim(UCase(wsLogSheet.Cells(i, "C").Value)) = "" Or _
                Trim(UCase(wsLogSheet.Cells(i, "C").Value)) = "DATA PENDING" Or _
                Trim(UCase(wsLogSheet.Cells(i, "C").Value)) = "N/A") Or _
               (Trim(UCase(wsLogSheet.Cells(i, "E").Value)) = "" Or _
                Trim(UCase(wsLogSheet.Cells(i, "E").Value)) = "DATA PENDING" Or _
                Trim(UCase(wsLogSheet.Cells(i, "E").Value)) = "OPEN") Then
                    
                Set foundCell = wsData.Range("A:A").Find(What:=caseID, LookIn:=xlValues, LookAt:=xlWhole)
                If Not foundCell Is Nothing Then
                    ' Update Owner (Column B) with the value from Data_Import (offset 1)
                    If Trim(wsLogSheet.Cells(i, "B").Value) <> Trim(foundCell.Offset(0, 1).Value) Then
                        wsLogSheet.Cells(i, "B").Value = foundCell.Offset(0, 1).Value
                        rowUpdated = True
                    End If

                    ' Update TimeCreated (Column C) from Data_Import (offset 2)
                    If Not IsDate(wsLogSheet.Cells(i, "C").Value) Or _
                       (Trim(UCase(CStr(wsLogSheet.Cells(i, "C").Value))) <> Trim(UCase(CStr(foundCell.Offset(0, 2).Value)))) Then
                        wsLogSheet.Cells(i, "C").Value = foundCell.Offset(0, 2).Value
                        rowUpdated = True
                    End If
                    
                    ' Update TimeClosed (Column E) from Data_Import (offset 4) if valid; else set to "Open"
                    If IsDate(foundCell.Offset(0, 4).Value) Then
                        If Not IsDate(wsLogSheet.Cells(i, "E").Value) Or _
                           (Trim(UCase(CStr(wsLogSheet.Cells(i, "E").Value))) = "OPEN") Then
                            wsLogSheet.Cells(i, "E").Value = foundCell.Offset(0, 4).Value
                            rowUpdated = True
                        End If
                    Else
                        wsLogSheet.Cells(i, "E").Value = "Open"
                    End If
                    
                    ' Recalculate MTTP (Column G) if TimeCreated (C) and QuickEntry Time (D) are dates
                    If IsDate(wsLogSheet.Cells(i, "C").Value) And IsDate(wsLogSheet.Cells(i, "D").Value) Then
                        wsLogSheet.Cells(i, "G").Value = FormatMinutes(DateDiff("n", wsLogSheet.Cells(i, "C").Value, wsLogSheet.Cells(i, "D").Value))
                    End If
                    
                    ' Update Late Note Status (Column H) based on pickup delay (from C to D)
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
                    
                    ' Recalculate MTTR (Column I) if both TimeCreated (C) and TimeClosed (E) are dates
                    If IsDate(wsLogSheet.Cells(i, "C").Value) And IsDate(wsLogSheet.Cells(i, "E").Value) Then
                        wsLogSheet.Cells(i, "I").Value = FormatMinutes(DateDiff("n", wsLogSheet.Cells(i, "C").Value, wsLogSheet.Cells(i, "E").Value))
                    Else
                        wsLogSheet.Cells(i, "I").Value = "Open"
                    End If
                    
                    ' Update Spike Detection (Column J) based on TimeCreated (Column C)
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
                    
                    If rowUpdated Then updatedCount = updatedCount + 1
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
