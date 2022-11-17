Attribute VB_Name = "colorCheck"
Option Explicit

 Sub color_check()
 
    Dim end_row, column, i As Integer
    Dim check_text As String
    
    column = 8
    end_row = Cells(Rows.Count, column).End(xlUp).Row
    check_text = "ïsçáäi"
    
    For i = 1 To end_row
       If Cells(i, column).Value = check_text Then
           ActiveSheet.Range(Cells(i, 1), Cells(i, column)).Interior.ColorIndex = 6
       End If
    Next
 
 End Sub
