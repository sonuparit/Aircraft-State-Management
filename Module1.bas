Attribute VB_Name = "Module1"
Sub Initialize_AC_State_Click()
'
'
'   To Copy AC State at Left
'
'
    Dim i As Integer
    Range("A7:P12").Select
    Selection.Copy
    Range("AE7:AT12").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A14:P19").Select
    Selection.Copy
    Range("AE14:AT19").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A21:P22").Select
    Selection.Copy
    Range("AE21:AT22").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D27:D31").Select
    Selection.Copy
    Range("AH27:AH31").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'
'
'   TO reset Daily Flg to ZERO
'
'
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "12:00:00 AM"
    Range("D7").Select
    Selection.Copy
    Range("E7").Select
    ActiveSheet.Paste
    Range("D8:E12").Select
    ActiveSheet.Paste
    Range("D14:E19").Select
    ActiveSheet.Paste
    Range("D21:E22").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'
'
'   For Setting G Run to ZERO
'
'
    Range("AV7").Select
    ActiveCell.FormulaR1C1 = "12:00:00 AM"
    Range("AV7").Select
    Selection.Copy
    Range("AV8:AV12").Select
    ActiveSheet.Paste
    Range("AV14:AV19").Select
    ActiveSheet.Paste
    Range("AV21:AV22").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'
'
'   To check the life of IDS below 12:00 Hrs
'
'
    For i = 28 To 33
        If Cells(i, 25).Value < 0.5 Then
        MsgBox (Cells(i, 22).Value & " IDS has only : " & Format(Cells(i, 25).Value, "hh:mm") & " Hrs left, give its FDR, Chip Data & Oil Sample for SOAP Analysis to HAL Det")
        End If
        If Cells(i, 29).Value < 0.5 Then
        MsgBox (Cells(i, 26).Value & " IDS has only : " & Format(Cells(i, 29).Value, "hh:mm") & " Hrs left, give its FDR, Chip Data & Oil Sample for SOAP Analysis to HAL Det")
        End If
    Next i
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
End Sub


