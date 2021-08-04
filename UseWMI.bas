Attribute VB_Name = "UseWMI"
Option Explicit

' Windows Management Instrumentation (WMI) を利用
' https://www.atmarkit.co.jp/icd/root/21/81148321.html
' Windowsのシステム管理用インターフェイス。
' エンタープライズ・レベルのシステム管理を目的として考案された標準仕様のWBEM（業界団体のDMTFが策定）に従って
' マイクロソフトがこれをWindows OS上に実装、拡張したもの。
' WMIが提供するインターフェイスを利用し､システム管理用コンポーネントにアクセスすることができる｡

' MACアドレスを取得

Public Sub getMacAddress()
  Dim objNetwork As Object 'Windowsの情報
  Dim strNetworkSql As String 'Windowsの情報取得文 保存変数
  Dim strMacAdr As String '取得したMACアドレス文字

  'Windowsの情報取得文 組み立て
  strNetworkSql = "SELECT * FROM Win32_NetworkAdapter WHERE MACAddress IS NOT NULL"

  'Windowsの情報取得文を使い情報を取得(1個目のみ)
  For Each objNetwork In GetObject("winmgmts:").ExecQuery(strNetworkSql)
    strMacAdr = objNetwork.MACAddress
  Next
  
  'メッセージボックスへMACアドレスを表示
  MsgBox (strMacAdr)

End Sub

Sub Sample()
  
  ''すべての自動実行を取得する
  Dim objWMIService, colStartupCommands, objStartupCommand
  Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
  Set colStartupCommands = objWMIService.ExecQuery("Select * from Win32_StartupCommand")
  
  For Each objStartupCommand In colStartupCommands
    Debug.Print "Command: " & objStartupCommand.Command
    Debug.Print "Description: " & objStartupCommand.Description
    Debug.Print "Location: " & objStartupCommand.Location
    Debug.Print "Name: " & objStartupCommand.Name
    Debug.Print "User: " & objStartupCommand.User
    Debug.Print ""
  Next
  Set colStartupCommands = Nothing
  Set objWMIService = Nothing
End Sub

