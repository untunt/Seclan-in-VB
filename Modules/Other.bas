Attribute VB_Name = "Other"
Sub ÑÝ³ö±àºÅ()
    Dim RowLast As Long
    Dim i As Long
    Dim v1 As Long
    Dim v2 As Long
    RowLast = ActiveSheet.UsedRange.Rows.Count
    v1 = 0
    For i = 2 To RowLast
        If InStr(1, Cells(i, 1), ">") Then
            v1 = v1 + 1
            Cells(i, 1) = Format(v1, "00") & " >"
            v2 = 0
        ElseIf InStr(1, Cells(i, 1), "]") Then
            v2 = v2 + 1
            Cells(i, 1) = Format(v1, "00") & "." & v2 & "]"
        End If
    Next i
End Sub
