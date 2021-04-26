Attribute VB_Name = "Module1"
Option Explicit

Sub ddd()
Dim w As Worksheet
For Each w In Worksheets
MsgBox "找到工作表:" & w.Name
Next
MsgBox "完成迴圈掃描"


End Sub
