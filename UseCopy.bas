Attribute VB_Name = "UseCopy"
Option Explicit

Sub ClipBordSet(CopySrc As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = CopySrc
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

Sub TestClipBord()
  Call ClipBordSet("test")
End Sub
