Attribute VB_Name = "SheetCopy"
Option Explicit

Sub SheetCopy()

'�ϐ���variant�ɂ��Ȃ��Ɛ���ɓ����Ȃ��B
Dim SheetName_Day As Variant
SheetName_Day = Day(Date)
Debug.Print SheetName_Day

'�V�[�g���͕ς��邱�ƁB
Worksheets("Sheet1").Copy after:=Worksheets(1)
ActiveSheet.name = 1
End Sub
