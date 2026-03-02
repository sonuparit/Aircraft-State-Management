Attribute VB_Name = "Module2"
Sub Daily_AC_Log_Entry_Click()
'
' Daily_AC_Log_Entry_Click Macro
'

'   For IA-1141

    ActiveWindow.ScrollColumn = 1
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1141").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Range("B7:Q7").Select
    Selection.Copy
    Sheets("IA-1141").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
        
'   For IA-1143
    
    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1143").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Range("B8:Q8").Select
    Selection.Copy
    Sheets("IA-1143").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
        
'   For IA-1145
    
    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1145").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Range("B9:Q9").Select
    Selection.Copy
    Sheets("IA-1145").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
        
'   For IA-1149
    
    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1149").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Range("B10:Q10").Select
    Selection.Copy
    Sheets("IA-1149").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
        
'   For IA-1173
        
    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1173").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Range("B11:Q11").Select
    Selection.Copy
    Sheets("IA-1173").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
              
        
'   For IA-1183
        
    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1183").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Range("B12:Q12").Select
    Selection.Copy
    Sheets("IA-1183").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
        
  
'   For IA-1147
    
    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1147").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Range("B14:Q14").Select
    Selection.Copy
    Sheets("IA-1147").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    
'   For IA-1171

    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1171").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Range("B15:Q15").Select
    Selection.Copy
    Sheets("IA-1171").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
        
'   For IA-1172
    
    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1172").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Range("B16:Q16").Select
    Selection.Copy
    Sheets("IA-1172").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
        
        
'   For IA-1174
    
    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1174").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Range("B17:Q17").Select
    Selection.Copy
    Sheets("IA-1174").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
        
'   For IA-1180
    
    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1180").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Range("B18:Q18").Select
    Selection.Copy
    Sheets("IA-1180").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
        
'   For IA-1148
    
    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1148").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Range("B19:Q19").Select
    Selection.Copy
    Sheets("IA-1148").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    
'   For IA-1144
    
    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1144").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Application.CutCopyMode = False
    Range("B21:Q21").Select
    Selection.Copy
    Sheets("IA-1144").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
        

'   For IA-1146
    
    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("IA-1146").Select
    Range("A4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    Sheets("NEW AC STATE").Select
    Range("B22:Q22").Select
    Selection.Copy
    Sheets("IA-1146").Select
    Range("C4").Select
    Selection.Insert Shift:=xlDown
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
    
    
'   Return to AC STATE

    Sheets("NEW AC STATE").Select
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollRow = 1
    Application.CutCopyMode = False
    
End Sub


