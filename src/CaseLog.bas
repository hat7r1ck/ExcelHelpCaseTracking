Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rng As Range
    ' Check if the change occurred in column F (Notes)
    If Not Intersect(Target, Me.Range("F:F")) Is Nothing Then
        Application.EnableEvents = False  ' Prevent re-triggering the event
        For Each rng In Intersect(Target, Me.Range("F:F"))
            ' If a note is now provided (non-empty), and Late Note Status (column H) says "NOTE REQUIRED", update it.
            If Len(Trim(rng.Value)) > 0 Then
                If UCase(Me.Cells(rng.Row, "H").Value) = "NOTE REQUIRED" Then
                    Me.Cells(rng.Row, "H").Value = "Note provided"
                    ' Remove any highlighting (set to no fill)
                    Me.Cells(rng.Row, "H").Interior.ColorIndex = xlNone
                End If
            End If
        Next rng
        Application.EnableEvents = True
    End If
End Sub

