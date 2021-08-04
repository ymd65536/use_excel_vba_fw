Attribute VB_Name = "UseListObject"
Option Explicit

' Excel ��ListObject���g���₷�����郂�W���[��

' ListObject�̍s����Ԃ�
Function getRowsCount(list_object As ListObject)
  getRowsCount = list_object.ListRows.Count
End Function

' ���X�g�I�u�W�F�N�g�P�ʂŃ\�[�g
Sub SortData(ws As Worksheet, column_name As String, _
Optional item_count As Integer = 1)

  Dim listObj As ListObject

  ws.Activate
  
  Set listObj = ws.ListObjects.Item(item_count)
  
  listObj.Range.AutoFilter
  listObj.Sort.SortFields.Clear
  listObj.Sort.SortFields.Add _
    Key:=ws.Range("D1"), _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:=xlSortNormal
    
  listObj.Sort.Header = xlYes
  listObj.Sort.Orientation = xlTopToBottom
  listObj.Sort.Apply

End Sub

' ���X�g�I�u�W�F�N�g�P�ʂŃt�B���^�[
Sub filterData(ws As Worksheet, column_name As String, _
Optional item_count As Integer = 1)
  ws.Activate
  ws.ListObjects.Item(1).Range.AutoFilter ws.ListObjects.Item(item_count).ListColumns.Item(column_name).Index, Criteria1:=">=" & Format(Now, "yyyy/m/d")
End Sub

' ���X�g�I�u�W�F�N�g�P�ʂŃZ����I��
Function returnColumnDataBody(ws As Worksheet, _
  column_name As String, _
  Optional item_count As Integer = 1) As Range
  ws.Activate
  Set returnColumnDataBody = ws.ListObjects.Item(item_count).ListColumns.Item(column_name).DataBodyRange
End Function

' ���X�g�I�u�W�F�N�g�P�ʂŃZ����I��
Sub selectColumnDataBody(ws As Worksheet, _
  column_name As String, _
  Optional item_count As Integer = 1)
  ws.Activate
  ws.ListObjects.Item(item_count).ListColumns.Item(column_name).DataBodyRange.Select
End Sub

' ���X�g�I�u�W�F�N�g�P�ʂŃZ�����R�s�[
Sub copyListObject(ws As Worksheet, _
  column_name As String, _
  Optional item_count As Integer = 1)
  Application.CutCopyMode = True
  ws.ListObjects.Item(item_count).ListColumns.Item(column_name).DataBodyRange.Copy
End Sub

' List�I�u�W�F�N�g�ɍ��킹�ăZ���ɓ\��t��
Sub pasteListObject(ws As Worksheet, _
  column_name As String, _
  Optional item_count As Integer = 1)
  
  ws.Activate
    ws.ListObjects.Item(item_count).ListColumns.Item(column_name).DataBodyRange.PasteSpecial xlPasteValues
End Sub

' �w��̃Z���ɓ\��t�� ���X�g�I�u�W�F�N�g�𗘗p
Sub pasteIdxListObject(ws As Worksheet, _
  column_name As String, _
  Optional item_count As Integer = 1, _
  Optional row_idx As Long = 2)
  
  ws.Activate
  ws.Cells(row_idx, ws.ListObjects.Item(item_count).ListColumns.Item(column_name).Index).PasteSpecial xlPasteValues
  
  Application.CutCopyMode = False
End Sub

' ListObject���N���A����
Sub ClearListObj(ws As Worksheet, rows_cnt As Long, Optional column_cnt As Integer = 2)
  If rows_cnt > 0 Then
    If rows_cnt = 1 Then
      rows_cnt = rows_cnt + 1
    End If
    ws.Range(CStr(rows_cnt) & ":" & column_cnt).Delete
  End If
End Sub

