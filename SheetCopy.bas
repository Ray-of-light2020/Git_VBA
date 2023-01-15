Attribute VB_Name = "SheetCopy"
Option Explicit

Sub SheetCopy()

'変数をvariantにしないと正常に動かない。
Dim SheetName_Day As Variant
SheetName_Day = Day(Date)
Debug.Print SheetName_Day

'シート名は変えること。
Worksheets("Sheet1").Copy after:=Worksheets(1)
ActiveSheet.name = 1
End Sub
