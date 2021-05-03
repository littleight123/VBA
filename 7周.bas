Attribute VB_Name = "Module1"
Option Explicit
Sub ¦X¨Ö()
Dim shtsIdx As Integer
For shtsIdx = 1 To Sheets.Count
Sheets(shtsIdx).Activate
   Dim k As Long
For k = 2 To 11 Step 3
    Dim rangStr As String
    rangStr = "A" & k & ":A" & k + 2
    Range(rangStr).Merge
Next
Next

End Sub

