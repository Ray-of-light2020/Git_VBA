Attribute VB_Name = "countMulti"
Option Explicit

Sub countMulti()

Dim i As Long
Dim j As Variant
Const N = 15
Dim x As Long: x = 3
Dim y As Long: y = 5
Dim myArray() As Long
ReDim myArray(N) As Long
Dim index As Long: index = 0

'x��y�̔{����z��Ɋi�[
For i = 1 To N
    If 0 = i Mod 3 Or 0 = i Mod 5 Then
     myArray(index) = i
     index = index + 1
    End If
Next

Debug.Print N & "�ȉ���" & x & "�̔{����" & y & "�̔{���̌��� " & index & "�ł�� "

'�z��̗v�f�����Ē�`
ReDim Preserve myArray(index - 1) As Long

For Each j In myArray
    Debug.Print (j)
Next


End Sub
