Attribute VB_Name = "Module1"
Option Explicit

Sub 動態合併2()
Dim sht As Integer
For sht = 1 To Sheets.Count
    Sheets(sht).Activate

    Application.DisplayAlerts = False '作業系統提醒文字，若沒有設定會依值提醒'
Dim i, j As Long '宣告i最後，j為長整數i為最後一列j為當前列索引'
Dim myrng As Range '宣告範圍變數'
'動態尋找A欄位有最後一列的列索引'
i = Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox "A欄位有資料最後一列索引" & i '說明用
For j = i To 2 Step -1 '從最後一列到第二列遞減，STEP-1為倒數
    Set myrng = Cells(j, "A") '目前範圍
    If myrng = myrng.Offset(-1, 0) Then '若目前的A欄位值和前一列相同
        myrng.Offset(-1, 0).Resize(2, 1).Merge '則需由下而上合併
        End If
Next
Next
Application.DisplayAlerts = True '重新開啟自動提醒

    
End Sub

