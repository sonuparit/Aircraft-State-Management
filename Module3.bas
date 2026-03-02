Attribute VB_Name = "Module3"
Sub AC_Serviceable_Days_Click()

'   For Date

    Sheets("S I R A DAYS").Select
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("N5").Select
    Sheets("NEW AC STATE").Select
    Range("D3:E3").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("A4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'   For IA-1141

    Sheets("NEW AC STATE").Select
    Range("B7").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("C4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'   For IA-1143

    Sheets("NEW AC STATE").Select
    Range("B8").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("D4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'   For IA-1145
        
    Sheets("NEW AC STATE").Select
    Range("B09").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("E4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'   For IA-1149

    Sheets("NEW AC STATE").Select
    Range("B10").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("F4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
                
                
'   For IA-1173
        
    Sheets("NEW AC STATE").Select
    Range("B11").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("G4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
'   For IA-1183
        
    Sheets("NEW AC STATE").Select
    Range("B12").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("H4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
'   For IA-1147
        
    Sheets("NEW AC STATE").Select
    Range("B14").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("I4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
'   For IA-1171
        
    Sheets("NEW AC STATE").Select
    Range("B15").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("J4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
'   For IA-1172
        
    Sheets("NEW AC STATE").Select
    Range("B16").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("K4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'   For IA-1174
        
    Sheets("NEW AC STATE").Select
    Range("B17").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("L4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False0
    
'   For IA-1180
    
    Application.CutCopyMode = False
    Sheets("NEW AC STATE").Select
    Range("B18").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("M4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
        
'   For IA-1148
        
    Sheets("NEW AC STATE").Select
    Range("B19").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("N4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
            
'   For IA-1144
        
    Sheets("NEW AC STATE").Select
    Range("B21").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("O4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
'   For IA-1146
        
    Sheets("NEW AC STATE").Select
    Range("B22").Select
    Selection.Copy
    Sheets("S I R A DAYS").Select
    Range("P4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'   Return to AC State
        
    Sheets("NEW AC STATE").Select
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollRow = 1
    Application.CutCopyMode = False
    
End Sub


