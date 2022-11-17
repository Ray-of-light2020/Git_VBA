Attribute VB_Name = "varidateCheck"
Option Explicit


Sub varidateCheck()
Dim rCell As Range
Const CHECK As String = "#.#"

For Each rCell In Selection
    If rCell.Value Like CHECK Then
    
    Else
       rCell.Interior.Color = RGB(240, 240, 20)
    End If
Next

End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'
   Selection.Interior.Color = xlNone
End Sub

