Attribute VB_Name = "UseAccess"
Option Explicit

' �Q�Ɛݒ�
'   Microsoft Scripting Runtime
'   Microsoft Excel 16.0 Object Library
'   Microsoft ActiveX Data Objects 7.1 Library
'
' �V�����֐������Ƃ���Access�Ə����n�߂�

Private oAcc                   As Access.Application
Private Const dbProviderStr    As String = "Microsoft.Ace.OLEDB.16.0"
Private ConStr                 As String

'
' �f�[�^�x�[�X�̃Z�b�g�A�b�v
'
Sub AccessDBSetUp(db As ADODB.Connection)
  Set db = New ADODB.Connection
  db.Provider = dbProviderStr
End Sub

'
' �f�[�^�x�[�X���J��
'
Sub AccessDataBaseOpen(db As ADODB.Connection, db_path As String)
  On Error GoTo Err_Exit
  db.Open db_path
  Exit Sub

Err_Exit:
  MsgBox Err.Description
End Sub
'
' ���R�[�h�Z�b�g���J��
'
Sub AccessOpenRecordSet(db As ADODB.Connection, _
  set_sql As String, _
  ByRef rs As ADODB.Recordset, Optional mode As Integer = adOpenStatic)
  rs.Open set_sql, db, mode
End Sub

'
' ���R�[�h�Z�b�g�̍s�����擾����B
'
Public Function AccessRecordsetCount(rs As ADODB.Recordset) As Integer
  If Not rs Is Nothing Then
    If rs.State = 1 Then
      AccessRecordsetCount = rs.RecordCount
    End If
  End If
  Exit Function
Err_Exit:
  If Err.Number <> 0 Then
    AccessRecordsetCount = 0
  End If
End Function

'
' �N�G�������s
'
Sub AccessOpenQuey(query_name As String, _
  list_file_pass As String, _
  file_name As String, _
  Optional db_path As String = "")
  
  If db_path = "" Then
    db_path = ThisWorkbook.Path & "\" & "test.accdb"
  End If

  'DB�I�[�v��
  Set oAcc = CreateObject("Access.Application")
  
  With oAcc
    .OpenCurrentDatabase (db_path)
    .Run query_name, list_file_pass, file_name
    .CloseCurrentDatabase
  End With
  Set oAcc = Nothing
End Sub
'
' Access����f�[�^�����o�� Excel�ɓ]�L
'
Sub CopyFromRecordset(ws As Worksheet, _
  rs As ADODB.Recordset, _
  Optional row_idx As Integer = 1, _
  Optional col_idx As Integer = 1)
  ws.Cells(row_idx, col_idx).CopyFromRecordset rs
End Sub
'
' ���R�[�h�Z�b�g�����
'
Sub AccessCloseRecordSet(ByRef rs As ADODB.Recordset)
  If Not rs Is Nothing Then
    rs.Close
    Set rs = Nothing
  Else
    Set rs = Nothing
  End If
End Sub
'
' �f�[�^�x�[�X�����
'
Sub AccessCloseDataBase(db As ADODB.Connection)
  If Not db Is Nothing Then
    db.Close
    Set db = Nothing
  Else
    Set db = Nothing
  End If
End Sub

'
' �O����Acess�f�[�^�x�[�X�ɂȂ�
'
'
' �p�X���[�h���f�[�^�x�[�X�̃I�[�v��
'
Function AccessConnectDataBase(db As ADODB.Connection, db_path As String, Optional password As String = "") As Boolean

  On Error GoTo Err_Exit
    
  ConStr = "Provider=" & dbProviderStr & ";" & _
  "Data Source=" & db_path & ";" & _
  "Jet OLEDB:Database Password=" & password & ";"
    
  Set db = New ADODB.Connection

  db.ConnectionString = ConStr
  db.Open
  
  If Not db Is Nothing Then
    AccessConnectDataBase = True
    Exit Function
  Else
    AccessConnectDataBase = False
    Exit Function
  End If
  
Exit Function
  
Err_Exit:
  
  MsgBox Err.Description & "(�G���[�ԍ��F" & Err.Number & ")"
  Call AccessCloseDataBase(db)
  AccessConnectDataBase = False
End Function

'
'
' �t�B�[���h�������݂��邩���m�F����B
'
'
Function AccessExistFieldName(rs As ADODB.Recordset, FieldNameStr As String) As Boolean

  Dim FieldObj As Field
  
  AccessExistFieldName = False

  If Not rs Is Nothing Then
    For Each FieldObj In rs.Fields
      If FieldObj.Name = FieldNameStr Then
        AccessExistFieldName = True
        Exit For
      End If
    Next
  End If
  
  Exit Function

Err_Exit:

  MsgBox Err.Description & "(AccessExistFieldName)"
  AccessExistFieldName = False

End Function
'
' �g�����U�N�V�������g��
'

'
' �g�����U�N�V�����J�n
'
Public Sub AccessBeginTrans(db As ADODB.Connection)
  On Error GoTo Err_Exit
  If Not db Is Nothing Then
    db.BeginTrans
  End If
  
  Exit Sub
Err_Exit:
  If Err.Number <> 0 Then
    
  End If
End Sub

'
' �g�����U�N�V�������R�~�b�g
'
Public Sub AccessComitTrans(db As ADODB.Connection)
  On Error GoTo Err_Exit
  db.CommitTrans
  
  Exit Sub
Err_Exit:
  If Err.Number <> 0 Then
  End If
End Sub

'
' �g�����U�N�V���������[���o�b�N
'
Public Sub AccessRollbackTrans(db As ADODB.Connection)

  On Error GoTo Err_Exit
  db.RollbackTrans
  
  Exit Sub
Err_Exit:
  If Err.Number <> 0 Then
  End If
End Sub

'
' 1���R�[�h�̓���̗񂾂��C������
'
Public Sub AccessEditRecordset(rs As ADODB.Recordset, Optional Col As String = "", Optional SetParam As String = "")
  If Col = "" Then
    MsgBox "����Col�F�񖼂͋󔒂ɂł��܂���! "
    GoTo Err_Exit
  End If
  If Not rs Is Nothing Then
    If rs.State = 1 Then
      rs(Col).Value = SetParam
    End If
  End If
  Exit Sub

Err_Exit:
  If Err.Number <> 0 Then
  End If
End Sub

'
' ���R�[�h�Z�b�g���X�V����B
'
Public Sub AccessUpdateRecordset(rs As ADODB.Recordset, Optional conform As Boolean = False)
  If conform = True Then
    MsgBox "�f�[�^���X�V���܂��B(rs.Update)"
  End If
  If Not rs Is Nothing Then
    If rs.State = 1 Then
      rs.Update
    End If
  End If
  Exit Sub

Err_Exit:
  If Err.Number <> 0 Then
  End If
End Sub

