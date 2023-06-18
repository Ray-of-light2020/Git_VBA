Attribute VB_Name = "weekend"
Option Explicit

 Sub week_end()
 
  Dim y As Long, m As Long
  Dim slast As Date '月末を格納
  Dim slastday As Long '月末日を格納
  Dim i As Long 'カウンター変数
  Dim cellValue As String 'セルの値を格納
  Dim thisMonth As Date '今月の各日を格納
  Dim dayOfweek As Long '各日の曜日を格納

  '初期値 今月を値で設定
  y = Year(Date)
  m = Month(Date)
  
  '月末を取得
  slast = DateSerial(y, m + 1, 0)
   '月末日を取得
  slastday = Format(slast, "d")
  
  For i = 1 To slastday
  '下記のセルの値で判定(方向やセル位置はここから変更)
    With Cells(1, i)
      cellValue = .Value
      thisMonth = y & "/" & m & "/" & cellValue
      dayOfweek = Weekday(thisMonth)
      
        If dayOfweek = 7 Then '土曜日なら
            .Interior.ColorIndex = 20
        ElseIf dayOfweek = 1 Then  '日曜なら
            .Interior.ColorIndex = 38
        End If
      End With
  Next i
   
  End Sub
