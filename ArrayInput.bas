Attribute VB_Name = "ArrayInput"
Option Explicit
'Array関数に格納する値をセルから取得
Sub ArrInput()

Const TargetColumn As Long = 1 '値を取得する列の定数
Dim base As Variant: base = Cells(1, TargetColumn).Value '初めの値を"A1"より取得
Dim inputArr As Long '二つ目以降の値を取得する変数
Dim i As Long 'ループ用のカウンター
Dim LastRow, xlLastRow As Long '最終行を取得するための変数

     xlLastRow = Cells(Rows.Count, 1).Row  'Excelの最終行を取得
    LastRow = Cells(xlLastRow, TargetColumn).End(xlUp).Row   'A列の最終行を取得
 
    For i = 2 To LastRow
        inputArr = Cells(i, 1).Value
        base = base & "," & inputArr
    Next i
    
Debug.Print base

End Sub

Sub makeArray()
 Dim v As Variant
 '下記（）にbase の結果をコピペ
 v = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
 
 Debug.Print (v(4))
 
 
End Sub

