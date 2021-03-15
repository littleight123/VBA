Attribute VB_Name = "Module1"
Sub 計算()
'方法一計算第一列到第五列的第五欄
Range("E1").Value = Range("A1").Value + Range("C1").Value
Range("E2").Value = Range("A1").Value - Range("C1").Value
Range("E3").Value = Range("A1").Value * Range("C1").Value
Range("E4").Value = Range("A1").Value / Range("C1").Value

'方法二計算第一列到第五列的第五欄
Cells(1, 5).Value = Cells(1, 1).Value + Cells(1, 3).Value
Cells(2, 5).Value = Cells(1, 1).Value - Cells(1, 3).Value
Cells(3, 5).Value = Cells(1, 1).Value * Cells(1, 3).Value
Cells(4, 5).Value = Cells(1, 1).Value / Cells(1, 3).Value

'方法三計算第一列到第五列的第五欄
Cells(1, "E").Value = Cells(1, "A").Value + Cells(1, "C").Value
Cells(2, "E").Value = Cells(1, "A").Value - Cells(1, "C").Value
Cells(3, "E").Value = Cells(1, "A").Value * Cells(1, "C").Value
Cells(4, "E").Value = Cells(1, "A").Value / Cells(1, "C").Value


End Sub
