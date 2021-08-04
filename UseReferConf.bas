Attribute VB_Name = "UseReferConf"
Option Explicit

'
' 参照設定を操作するモジュール
'
'利用できる参照設定をすべて見る
Sub Sample1()
    Dim Ref, buf As String
    For Each Ref In ActiveWorkbook.VBProject.References
        buf = buf & Ref.Name & vbTab & Ref.Description & vbCrLf
    Next Ref
    MsgBox buf
End Sub

'参照設定を追加する
Sub Sample2()
    Const RefFile As String = "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
    ActiveWorkbook.VBProject.References.AddFromFile RefFile
End Sub

'参照設定を外す
Sub Sample3()
  Dim Ref
  With ActiveWorkbook.VBProject
    For Each Ref In ActiveWorkbook.VBProject.References
      If Ref.Description = "Microsoft DAO 3.6 Object Library" Then
        .References.Remove Ref
      End If
    Next Ref
  End With
End Sub

'特定の参照設定が利用できるかを確認する
Sub Sample4()
  Dim Ref, buf As String
  For Each Ref In ActiveWorkbook.VBProject.References
    buf = buf & Ref.IsBroken & vbCrLf
  Next Ref
  MsgBox buf
End Sub
