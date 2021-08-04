Attribute VB_Name = "UseExcel"
Option Explicit

'
' Excelを使いやすくするモジュール
'

'
' 色をセット
'
Sub SetCellInteriorColor(rng As Range, color_num As Long)
  rng.Interior.Color = color_num
End Sub

'
' シートを見つけて返す
'
Function ReturnWorkSheet(wb As Workbook, sheet_name As String) As Worksheet
  Dim ws As Worksheet
  For Each ws In wb.Worksheets
    If sheet_name = ws.Name Then
      Set ReturnWorkSheet = ws
      Exit Function
    End If
  Next
  Set ReturnWorkSheet = Nothing
End Function

'
' ブックを見つけて返す
'
Function ReturnWorkbook(book_name As String) As Workbook
  Dim wb As Workbook
  For Each wb In Workbooks
    If book_name = wb.Name Then
      Set ReturnWorkbook = wb
      Exit Function
    End If
  Next
  Set ReturnWorkbook = Nothing
End Function

'
' データコピー
'
Sub CopyRng(SrcRng As Range, Dest As Range)
  SrcRng.Copy
  Dest.PasteSpecial xlPasteValues
  Application.CutCopyMode = False
End Sub

'
' WorksheetFunctionのVlookupを実行
'
Function WorksheetVLookUp(rng As Range, _
  ws As Worksheet, _
  Search As String, _
  resultInt As Integer) As String
  WorksheetVLookUp = WorksheetFunction.VLookup(rng, ws.Range(Search), resultInt, False)
End Function
'
' WorksheetFunctionのCountIfを実行
'
Function WorksheetCountIf(ws As Worksheet, _
  rng As String, _
  Search As String) As Integer
  
  WorksheetCountIf = WorksheetFunction.CountIf(ws.Range(rng), Search)
End Function
'
' End メソッドを利用した行数カウント
'
Function EndXlDownRowCount(ws As Worksheet, _
  RowIndex As Integer, _
  ColumIndex As Integer) As Integer
  EndXlDownRowCount = ws.Cells(RowIndex, ColumIndex).End(xlDown).Row
End Function

Sub SortCells(ws As Worksheet, SortRng As String, key1Rng As String, _
          Optional SortOrder As Integer = xlAscending)
  With ws
    .Range(SortRng).Sort Range(key1Rng), SortOrder
  End With
End Sub
'
' ハイパーリンクの追加
'
Sub AddHyperLink(ws As Worksheet, _
  url As String, _
  RowIndex As Long, _
  ColumIndex As Long)
  
  Dim LinkObj As Hyperlink
  
  With ws
    Set LinkObj = .Hyperlinks.Add(anchor:=.Cells(RowIndex, ColumIndex), Address:=url)
  End With
End Sub
'
' ハイパーリンクの削除
'
Sub DeleteHyperLink(wk As Workbook, _
  ShName As String, _
  RowIndex As Integer, _
  ColumIndex As Integer)
  
  With wk.Worksheets(ShName).Cells(RowIndex, ColumIndex)
    .Hyperlinks.Delete
  End With
End Sub
'ハイパーリンクの取得
Function GetHyperLink(wk As Workbook, _
  ShName As String, _
  RowIndex As Integer, _
  ColumIndex As Integer) As String
  
  With wk.Worksheets(ShName).Cells(RowIndex, ColumIndex)
    If .Hyperlinks.Count > 0 Then
      GetHyperLink = .Hyperlinks.Item(1).Address
    Else
      GetHyperLink = ""
    End If
  End With
End Function
'ハイパーリンクの有無確認
Function ExistHyperLink(wk As Workbook, _
  ShName As String, _
  RowIndex As Integer, _
  ColumIndex As Integer) As Boolean
  
  With wk.Worksheets(ShName).Cells(RowIndex, ColumIndex)
    If .Hyperlinks.Count > 0 Then
      ExistHyperLink = True
    Else
      ExistHyperLink = False
    End If
  End With
End Function
' リンク名とリンク先をFileSystemObjectから作成
Sub setHyperLinkByFilesystemObj(rng As Range, link_name As String, url As String)
  Dim LinkObj As Hyperlink
  rng.Value = link_name
  Set LinkObj = rng.Hyperlinks.Add(anchor:=rng, Address:=url)
End Sub
