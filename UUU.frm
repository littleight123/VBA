VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UUU 
   Caption         =   "UserForm1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "UUU.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UUU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnQuery_Click()
Dim fruit, L As String
Dim rown As Integer
fruit = txbFruitName.Text

For rown = 2 To 7
If (Cells(rown, "A").Value = fruit) Then
    lblFruitResult.Caption = Cells(rown, "B").Value
    If (Cells(rown, "C").Value = "Y") Then
        L = Cells(rown, "D").Value
       MsgBox (L & "買的到哦")
End If
End If

Next
End Sub




Private Sub CommandButton1_Click()

Dim FruitName, K As String
Dim rown As Integer
FruitName = txbFruitName.Text
For rown = 2 To 7
If (Cells(rown, "A").Value = FruitName) Then
    K = Cells(rown, "C").Value
    lblfresult.Caption = K
End If
Next
End Sub

