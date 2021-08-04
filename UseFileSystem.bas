Attribute VB_Name = "UseFileSystem"
Option Explicit

' �Q�Ɛݒ�
'   Microsoft Scripting Runtime
'FileSystem�I�u�W�F�N�g���g�����t�H���_�쐬

Sub MkdirFileSystem(FilePath As String)
  Dim pFileObj       As New FileSystemObject
  If pFileObj.FolderExists(FilePath) = False Then
    MkDir FilePath
  End If
  Set pFileObj = Nothing
End Sub

'FileSystem�I�u�W�F�N�g���g�����t�@�C���R�s�[
Sub FileCopyFileSystem(SrcFilePath As String, _
  DestFilePath As String, _
  Optional DeleteSrcFile As Boolean = True)

  Dim pFileObj       As New FileSystemObject
  If DeleteSrcFile Then
    pFileObj.DeleteFile DestFilePath, Force:=True
  End If
  pFileObj.CopyFile SrcFilePath, DestFilePath
  Set pFileObj = Nothing
End Sub

Sub FolderCopyFileSystem(SrcFilePath As String, DestFilePath As String)
  Dim pFileObj       As New FileSystemObject
  Call pFileObj.CopyFolder(SrcFilePath, DestFilePath, True)
  Set pFileObj = Nothing
End Sub


'FileSystem�I�u�W�F�N�g���g�����t�@�C���ړ�

Sub FileMoveFileSystem(SrcFilePath As String, DestFilePath As String)
  Dim pFileObj As New FileSystemObject
  If pFileObj.FileExists(SrcFilePath) Then
    pFileObj.MoveFile SrcFilePath, DestFilePath
  End If
  Set pFileObj = Nothing
End Sub

'FileSystem�I�u�W�F�N�g���g�����t�@�C���폜
Sub FileDeleteFileSystem(SrcFileName As String)
  Dim pFileObj As New FileSystemObject
  
  If pFileObj.FileExists(SrcFileName) = True Then
    pFileObj.DeleteFile SrcFileName
  End If
  Set pFileObj = Nothing
End Sub

'
' �w��̃t�H���_�z���ɂ���t�H���_���I�u�W�F�N�g�Ƃ��ĕԂ�
'
Function getFoldersObj(folder_path As String) As Folders
  Dim pFileObj As New FileSystemObject
  If pFileObj.FolderExists(folder_path) Then
    Set getFoldersObj = pFileObj.GetFolder(folder_path).SubFolders
  Else
    Set getFoldersObj = Nothing
  End If
End Function


'
' �t�H���_�̂���Ȃ�
'
Function IsExistFolder(folder_path As String) As Boolean
  Dim pFileObj As New FileSystemObject
  If pFileObj.FolderExists(folder_path) Then
    IsExistFolder = True
  Else
    IsExistFolder = False
  End If
End Function

