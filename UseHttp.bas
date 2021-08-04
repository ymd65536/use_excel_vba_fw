Attribute VB_Name = "UseHttp"
Option Explicit

' 参照設定 Microsoft XML, v6.0」

Sub xmlhttprequest_test()

  Dim Req As XMLHTTP60
  
  Set Req = New XMLHTTP60
  
  Req.Open "GET", "http://example"
  Req.send
  
  Do While Req.ReadyState < 4
      DoEvents
  Loop
  
  Debug.Print Req.responseStream
  
  Set Req = Nothing

End Sub

