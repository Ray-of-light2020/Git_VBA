Attribute VB_Name = "Module1"
Option Explicit

Sub nextInput(startDay, endDay)

Dim i As Long, lastColumn As Long, getDay As Long, betweenDay As Long, checkDay As Long

betweenDay = endDay - startDay + 1
checkDay = startDay

For i = 1 To betweenDay
    lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
    '��ԏ��߂̃Z������ł�A��ɂ͓��͂���Ȃ��B
    Cells(1, lastColumn + 1).Value = checkDay
    checkDay = checkDay + 1
Next i

End Sub

Sub start()
  UserForm1.Show
End Sub
