Attribute VB_Name = "SheetCopy"
Option Explicit

Sub sheetCopy()
 Dim nameDay As Variant
 
 '��ʂ̍X�V���~�߂�
 Application.ScreenUpdating = False
 nameDay = InputBox("�V�[�g������t�ō쐬���܂��B", "������" & Date & "�ł��B", day(Date))
 
 '���͒l�̊m�F
 If nameDay = "" Then '�L�����Z���������ꂽ��I��
    Exit Sub
 ElseIf IsNumeric(nameDay) = True Then
     Call Sheets("Sheet1").Copy(after:=Sheets(Worksheets.Count))
     ActiveSheet.name = nameDay
 Else
    MsgBox ("���ɂ�(���l)����͂��ĉ������B")
    Call sheetCopy
 End If
  Application.ScreenUpdating = True
End Sub
