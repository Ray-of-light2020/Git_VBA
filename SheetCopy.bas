Attribute VB_Name = "SheetCopy"
Option Explicit

Sub sheetCopy()
 Dim nameDay As Variant
 
 '画面の更新を止める
 Application.ScreenUpdating = False
 nameDay = InputBox("シート名を日付で作成します。", "今日は" & Date & "です。", day(Date))
 
 '入力値の確認
 If nameDay = "" Then 'キャンセルが押されたら終了
    Exit Sub
 ElseIf IsNumeric(nameDay) = True Then
     Call Sheets("Sheet1").Copy(after:=Sheets(Worksheets.Count))
     ActiveSheet.name = nameDay
 Else
    MsgBox ("日にち(数値)を入力して下さい。")
    Call sheetCopy
 End If
  Application.ScreenUpdating = True
End Sub
