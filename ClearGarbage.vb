'Private Sub ClearGarbage()
'  Dim lastline As Long: lastline = Worksheets(シート名).Cells(Rows.Count, 13).End(xlUp).Row
'
'  If lastline < getLowerEnd Then
'    Rows(lastline + 1 & ":" & getLowerEnd).Delete
'  End If
'End Sub

Private Sub ClearGarbage()
  Dim lastline As Long: lastline = Worksheets(シート名).Cells(Rows.Count, 13).End(xlUp).Row
  
  If lastline < getLowerEnd Then
    Rows(lastline + 1 & ":" & getLowerEnd).Delete
  End If
End Sub

Private Function getLowerEnd() As Long
  Const LASTCOL As Long = 20
  Dim lowerend As Long: lowerend = 1
  
  Dim i As Long
  For i = 1 To LASTCOL
    Dim tmp As Long: tmp = Worksheets(シート名).Cells(Rows.Count, i).End(xlUp).Row
    If lowerend < tmp Then
      lowerend = tmp
    End If
  Next
  
  getLowerEnd = lowerend
End Function