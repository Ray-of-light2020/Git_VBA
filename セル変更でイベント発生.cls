VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'指定したセル範囲の変更があった場合にプログラムを実行する。
Private Sub Worksheet_Change(ByVal Target As Range)

Dim tcell As Range
Dim makeSheet As Long

For Each tcell In Target '複数セルをひとつづつ判定
 If Not Application.Intersect(tcell, Range("A1:A5,C1:C5")) Is Nothing Then 'セル範囲に含まれるもののみ実行
  Debug.Print tcell.Address
  makeSheet = MsgBox("シートを作成しますか？", vbOKCancel)
    If makeSheet = 1 Then
        
    Else
        Exit Sub
    End If
 End If
Next
 

End Sub

