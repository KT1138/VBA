Option Explicit

Private Sub SortTemp2()
  Application.ScreenUpdating = False
  CopyTempToTemp2
  ItemSort
  Application.ScreenUpdating = True
End Sub

' 事前にTEMP2シートを作成しておくこと
Private Sub CopyTempToTemp2()
  Const SHEETNAME1 As String = "TEMP"
  Const SHEETNAME2 As String = "TEMP2"
  Const ORDERCOL As Long = 7   '「注文番号」列
  Const LASTCOL As Long = 30   '左端の列(会社で調べること)
  
  Worksheets(SHEETNAME1).Select
  Worksheets("SHEETNAME1").Range(Cells(1, 1), Cells(Cells(Rows.Count, ORDERCOL). _
    End(xlUp).Row, LASTCOL)).Copy Destination:=Worksheets(SHEETNAME2).Cells(1, 1)
End Sub

Private Sub ItemSort()
  Const SHEETNAME As String = "TEMP2"   'TEMPとは別にTEMP2シートを用意。
  'HOUSOU_PRINT と KONPOU_PRINT の最初のほうにある定数
  'SHEET_NAME2 には TEMP ではなく TEMP2 を設定すること
  Const ITEMCOL As Long = 3   '「品番」列
  Const ORDERCOL As Long = 7   '「注文番号」列
  Const NAMECOL As Long = 18   '「段取用」列
  Const FIRSTLINE As Long = 1
  
  'Worksheets(SHEETNAME).Select
  Dim ws As Worksheet: Set ws = Worksheets(SHEETNAME)
  Dim lastline As Long: lastline = _
    ws.Cells(Rows.Count, ORDERCOL).End(xlUp).Row
  
  Dim top As Long: top = 2
  Dim bot As Long
  For bot = top To lastline
    If ws.Cells(top, NAMECOL).Value <> ws.Cells(bot, NAMECOL) Or _
      bot = lastline Then
      ' top - bot 区間の行に対し、品番列をキーにして昇り順にソート
      ws.Rows(top & ":" & bot - 1).Sort key1:=ws.Cells(FIRSTLINE, ITEMCOL), _
        order1:=xlAscending ', Header:=xlYes
      top = bot
    End If
  Next bot
  
  Set ws = Nothing
End Sub
