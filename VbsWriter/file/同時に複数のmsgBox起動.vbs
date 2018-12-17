Set HtmA = CreateObject("htmlfile")
HtmA.ParentWindow.SetTimeout GetRef("ShowMsgA"), 0

Set HtmB = CreateObject("htmlfile")
HtmB.ParentWindow.SetTimeout GetRef("ShowMsgB"), 50

Set HtmC = CreateObject("htmlfile")
HtmC.ParentWindow.SetTimeout GetRef("ShowMsgC"), 100

Sub ShowMsgA
MsgBox "a"
End Sub

Sub ShowMsgB
MsgBox "b"
End Sub

Sub ShowMsgC
MsgBox "c"
End Sub

WScript.Sleep 100