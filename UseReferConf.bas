Attribute VB_Name = "UseReferConf"
Option Explicit

'
' �Q�Ɛݒ�𑀍삷�郂�W���[��
'
'���p�ł���Q�Ɛݒ�����ׂČ���
Sub Sample1()
    Dim Ref, buf As String
    For Each Ref In ActiveWorkbook.VBProject.References
        buf = buf & Ref.Name & vbTab & Ref.Description & vbCrLf
    Next Ref
    MsgBox buf
End Sub

'�Q�Ɛݒ��ǉ�����
Sub Sample2()
    Const RefFile As String = "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
    ActiveWorkbook.VBProject.References.AddFromFile RefFile
End Sub

'�Q�Ɛݒ���O��
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

'����̎Q�Ɛݒ肪���p�ł��邩���m�F����
Sub Sample4()
  Dim Ref, buf As String
  For Each Ref In ActiveWorkbook.VBProject.References
    buf = buf & Ref.IsBroken & vbCrLf
  Next Ref
  MsgBox buf
End Sub
