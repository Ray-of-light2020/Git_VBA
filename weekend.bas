Attribute VB_Name = "weekend"
Option Explicit

 Sub week_end()
  
 ' Dim week As Integer
  Dim y, m, week As Date
   
  y = Year(Date)
  m = Month(Date)

  
  For i = 1 To 31
  '�ϐ��Ɋi�[����ƃI�[�o�[�t���[����
  'ymd = y & "/" & m & "/" & i
  week = Weekday(y & "/" & m & "/" & i)
  'week = Weekday("2022/10/" & i)
 
    If week = 7 Then
        ActiveSheet.Cells(i, 1).Interior.ColorIndex = 20
    ElseIf week = 1 Then
        ActiveSheet.Cells(i, 1).Interior.ColorIndex = 38
    End If
  Next
   
  End Sub
