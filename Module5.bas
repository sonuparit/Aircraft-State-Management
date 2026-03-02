Attribute VB_Name = "Module5"
Sub Comp_Due_Check()
'
'
'   For monitoring of Comp life below 30:00 Hrs or
'   Below 30 Days in Cal basis
'
'
    Dim ac(1 To 11) As String
    ac(1) = "IA1141"
    ac(2) = "IA1143"
    ac(3) = "IA1145"
    ac(4) = "IA1146"
    ac(5) = "IA1147"
    ac(6) = "IA1149"
    ac(7) = "IA1171"
    ac(8) = "IA1172"
    ac(9) = "IA1173"
    ac(10) = "IA1174"
    ac(11) = "IA1180"
'    ac(12) = "IA-1143"
    Dim i, j As Integer
    Sheets("30H-30D Comp").Select
    Rows("5:5").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("6:1000").Select
    Selection.Delete Shift:=xlUp
    For j = 1 To 11
        Sheets(ac(j)).Select
        For i = 7 To 200
            Sheets(ac(j)).Select
            If IsEmpty(Cells(i, 13).Value) = True Then
                '   Do Nothing, but accend the value of i
            Else
            
            
                '   For Hrs
                
                
                If Cells(i, 13).Value < 1.25 Then
                    Sheets("30H-30D Comp").Select
                    Rows("5:5").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrBelow
                    Sheets(ac(j)).Select
                    Range(Cells(i, 3), Cells(i, 15)).Select
                    Selection.Copy
                    Sheets("30H-30D Comp").Select
                    Range("A5").Select
                    ActiveSheet.Paste
                    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
                        xlNone, SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                    GoTo Do_not_repeat  '    Skip Days for current row
                End If
                
                
                '   For Days
                
                
                If Cells(i, 14).Value < 30 Then
                    Sheets("30H-30D Comp").Select
                    Rows("5:5").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrBelow
                    Sheets(ac(j)).Select
                    Range(Cells(i, 3), Cells(i, 15)).Select
                    Selection.Copy
                    Sheets("30H-30D Comp").Select
                    Range("A5").Select
                    ActiveSheet.Paste
                    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
                        xlNone, SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                End If
            End If
Do_not_repeat:
        Next i
        Sheets("30H-30D Comp").Select
        Rows("5:5").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrBelow
        Sheets(ac(j)).Select
        Range("N3:O3").Select
        Selection.Copy
        Sheets("30H-30D Comp").Select
        Range("A5").Select
        ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
        Range("A5:M5").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        Selection.Merge
        Application.CutCopyMode = False
    Next j
End Sub
