VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "�J�n���ƏI�����̓���.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

    If IsNumeric(TextBox1) = True And IsNumeric(TextBox2.Text) = True Then
        Call nextInput(TextBox1.Text, TextBox2.Text)
        Unload UserForm1
    Else
        MsgBox ("��������͂��ĉ�����")
    End If
    
End Sub

Private Sub UserForm_Initialize()
  TextBox1.Text = Day(Date)
  TextBox2.Text = Day(Date)
End Sub
