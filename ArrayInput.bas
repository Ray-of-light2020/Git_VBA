Attribute VB_Name = "ArrayInput"
Option Explicit
'Array�֐��Ɋi�[����l���Z������擾
Sub ArrInput()

Const TargetColumn As Long = 1 '�l���擾�����̒萔
Dim base As Variant: base = Cells(1, TargetColumn).Value '���߂̒l��"A1"���擾
Dim inputArr As Long '��ڈȍ~�̒l���擾����ϐ�
Dim i As Long '���[�v�p�̃J�E���^�[
Dim LastRow, xlLastRow As Long '�ŏI�s���擾���邽�߂̕ϐ�

     xlLastRow = Cells(Rows.Count, 1).Row  'Excel�̍ŏI�s���擾
    LastRow = Cells(xlLastRow, TargetColumn).End(xlUp).Row   'A��̍ŏI�s���擾
 
    For i = 2 To LastRow
        inputArr = Cells(i, 1).Value
        base = base & "," & inputArr
    Next i
    
Debug.Print base

End Sub

Sub makeArray()
 Dim v As Variant
 '���L�i�j��base �̌��ʂ��R�s�y
 v = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
 
 Debug.Print (v(4))
 
 
End Sub

