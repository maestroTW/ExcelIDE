Sub vba_thisworkbook()
    glob = ""
    For i = 2 To 1000
        Worksheets(1).Cells(i, 1).Select

        If Application.Selection.Value = "<END>" Then
            Application.Selection.Offset(1, 0).Select
            Application.Selection.Value = glob
            Exit For
        ElseIf IsEmpty(Application.Selection.Value) Then
            For j = 1 To 100
                Worksheets(1).Cells(i, j).Select
                If IsEmpty(Application.Selection.Value) Then
                    glob = glob + vbTab
                Else
                    Exit For
                End If
            Next
        End If
        
        glob = glob + Application.Selection.Value
        glob = glob + vbCrLf
    Next
End Sub

