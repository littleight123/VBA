Attribute VB_Name = "Module1"
Option Explicit

Sub �ʺA�X��2()
Dim sht As Integer
For sht = 1 To Sheets.Count
    Sheets(sht).Activate

    Application.DisplayAlerts = False '�@�~�t�δ�����r�A�Y�S���]�w�|�̭ȴ���'
Dim i, j As Long '�ŧii�̫�Aj�������i���̫�@�Cj����e�C����'
Dim myrng As Range '�ŧi�d���ܼ�'
'�ʺA�M��A��즳�̫�@�C���C����'
i = Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox "A��즳��Ƴ̫�@�C����" & i '������
For j = i To 2 Step -1 '�q�̫�@�C��ĤG�C����ASTEP-1���˼�
    Set myrng = Cells(j, "A") '�ثe�d��
    If myrng = myrng.Offset(-1, 0) Then '�Y�ثe��A���ȩM�e�@�C�ۦP
        myrng.Offset(-1, 0).Resize(2, 1).Merge '�h�ݥѤU�ӤW�X��
        End If
Next
Next
Application.DisplayAlerts = True '���s�}�Ҧ۰ʴ���

    
End Sub

