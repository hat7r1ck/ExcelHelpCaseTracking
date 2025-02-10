Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanExit
    Dim cell As Range
    ' Check if any changes occur in column F (Notes)
    If Not Intersect(Target, Me.Range("F:F")) Is Nothing Then
        Application.EnableEvents = False  ' Disable events to prevent re-triggering
        For Each cell In Intersect(Target, Me.Range("F:F"))
            Debug.Print "Row " & cell.Row & " in column F changed to: " & cell.Value
            ' If the cell in column F is non-empty, update the corresponding cell in column H to "Note provided"
            If Len(Trim(cell.Value)) > 0 Then
                Me.Cells(cell.Row, "H").Value = "Note provided"
                Debug.Print "Row " & cell.Row & " in column H updated to 'Note provided'."
            End If
        Next cell
    End If
CleanExit:
    Application.EnableEvents = True  ' Ensure events are re-enabled
    If Err.Number <> 0 Then
        MsgBox "An error occurred: " & Err.Description, vbExclamation, "Error"
    End If
End Sub

