Attribute VB_Name = "Module1"
Global INorOUT As String
Global ShowPic As String



Sub Toggle()

On Error GoTo TryToFix

With MainWindow
Retry:
If INorOUT = "IN" Then
ShowPic = "Expand"
    'GOES IN
    Do
    DoEvents
    .Left = .Left + 100
    Loop Until .Left >= (Screen.Width - 200)
    INorOUT = "OUT"
    .Left = Screen.Width - 160 ' Sets the form off by 20 pixels, so when you move your mouse over to the form, its only over the expander button.
Else
ShowPic = "Reverse"
    'GOES OUT
    Do
    DoEvents
    .Left = .Left - 100
    Loop Until .Left < (Screen.Width - .Width + 100)
    INorOUT = "IN"

End If

Exit Sub

TryToFix:

' This is just in case you are in Maximized mode. You can't change certain options in this mode
.WindowState = vbNormal
DoEvents
GoTo Retry
End With



End Sub
