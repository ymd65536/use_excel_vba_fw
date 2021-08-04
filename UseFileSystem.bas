Attribute VB_Name = "UseFileSystem"
Option Explicit

' 参照設定
'   Microsoft Scripting Runtime
'FileSystemオブジェクトを使ったフォルダ作成

Sub MkdirFileSystem(FilePath As String)
  Dim pFileObj       As New FileSystemObject
  If pFileObj.FolderExists(FilePath) = False Then
    MkDir FilePath
  End If
  Set pFileObj = Nothing
End Sub

'FileSystemオブジェクトを使ったファイルコピー
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


'FileSystemオブジェクトを使ったファイル移動

Sub FileMoveFileSystem(SrcFilePath As String, DestFilePath As String)
  Dim pFileObj As New FileSystemObject
  If pFileObj.FileExists(SrcFilePath) Then
    pFileObj.MoveFile SrcFilePath, DestFilePath
  End If
  Set pFileObj = Nothing
End Sub

'FileSystemオブジェクトを使ったファイル削除
Sub FileDeleteFileSystem(SrcFileName As String)
  Dim pFileObj As New FileSystemObject
  
  If pFileObj.FileExists(SrcFileName) = True Then
    pFileObj.DeleteFile SrcFileName
  End If
  Set pFileObj = Nothing
End Sub

'
' 指定のフォルダ配下にあるフォルダをオブジェクトとして返す
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
' フォルダのあるなし
'
Function IsExistFolder(folder_path As String) As Boolean
  Dim pFileObj As New FileSystemObject
  If pFileObj.FolderExists(folder_path) Then
    IsExistFolder = True
  Else
    IsExistFolder = False
  End If
End Function

