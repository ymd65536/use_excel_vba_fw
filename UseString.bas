Attribute VB_Name = "UseString"
Option Explicit

' �C�C�����̕����񏈗����W���[��

Function StringDateYYYYMMDD(date_str As String)
  
  Dim spl_str() As String
  
  ' �X���b�V�������邩
  If InStr(date_str, "/") = 5 Then
    spl_str = Split(date_str, "/")
    If UBound(spl_str) = 2 Then
      StringDateYYYYMMDD = spl_str(0) & spl_str(1) & spl_str(2)
    End If
    If UBound(spl_str) = 3 Then
      StringDateYYYYMMDD = spl_str(0) & spl_str(1) & spl_str(2) & spl_str(3)
    End If
  Else
    StringDateYYYYMMDD = ""
  End If
End Function
