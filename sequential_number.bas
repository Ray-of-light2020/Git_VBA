Attribute VB_Name = "sequential_number"
Option Explicit

Sub sequential_number()

Dim i, j As Long

For i = 1 To 20 Step 3
    j = i + 2
    Debug.Print i & "Å`" & j
    Call sheet_copy(i, j)
Next

End Sub

Sub sheet_copy(ByVal i As Long, ByVal j As Long)

Worksheets("Sheet1").Copy after:=Worksheets(Worksheets.Count)
ActiveSheet.Name = i & "Å`" & j

End Sub
