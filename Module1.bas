Attribute VB_Name = "Module1"

Sub today()
  
 Dim start_work, Leaving_work As Date
 Dim myWSheet As Worksheet
 Dim today_ As Variant
 Dim search_column As Long
 
 start_work = "8:00"
 Leaving_work = "17:00" 
 
 serach_column = 1
 today_ = Application.InputBox(prompt:="以下の日にちに入力します。", Default:=Date)
 day_set = serch(serach_column, today_)
  
 ActiveSheet.Cells(day_set, serach_column).Interior.Color = vbYellow
 Application.GoTo Cells(day_set, serach_column), True
 
  If Cells(day_set, 2).Value = "" And Cells(day_set, 3).Value = "" Then
    With ActiveSheet.Cells(day_set, 2)
       .Value = start_work
       .NumberFormatLocal = "[h]:mm"
     End With
     With ActiveSheet.Cells(day_set, 3)
      .Value = Leaving_work
      .NumberFormatLocal = "[h]:mm"
    End With
        MsgBox "入力値の確認をお願いします。"
    
    Else
        MsgBox today_ & "日は入力済みです｡"
 End If
  ActiveSheet.Cells(day_set, 1).Interior.Color = xlNone
  
  
  End Sub
  
  Function serch(ByVal serch_column As Integer, ByVal today_ As Variant)
  
  Dim endcell As Integer
  Dim today As Integer
  
  
  today = Day(today_)
  endcell = Cells(Rows.Count, 1).End(xlUp).Row 
  
  serch = WorksheetFunction.Match(today, Range(Cells(1, serch_column), Cells(endcell, serch_column)), 0) 
  
  End Function
  
 
