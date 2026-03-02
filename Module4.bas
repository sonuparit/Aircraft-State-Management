Attribute VB_Name = "Module4"
Sub Gen_Soft_Copy_of_State_Click()
'
'   To gen soft copy of State Only Without Formulas
'

    Dim folder As String, filename As String
    Dim n As String
    n = Range("D3").Value
    folder = ThisWorkbook.Path & "\AC_States\" & Format(Date, "yyyy") & "\" & Format(Date, "mmmm") & "\"
    If Dir(folder, vbDirectory) = "" Then
        If Dir(ThisWorkbook.Path & "\AC_States\", vbDirectory) = "" Then MkDir (ThisWorkbook.Path & "\AC_States\")
        If Dir(ThisWorkbook.Path & "\AC_States\" & Format(Date, "yyyy"), vbDirectory) = "" Then MkDir (ThisWorkbook.Path & "\AC_States\" & Format(Date, "yyyy"))
        If Dir(ThisWorkbook.Path & "\AC_States\" & Format(Date, "yyyy") & "\" & Format(Date, "mmmm") & "\", vbDirectory) = "" Then MkDir (ThisWorkbook.Path & "\AC_States\" & Format(Date, "yyyy") & "\" & Format(Date, "mmmm") & "\")
        End If
    Sheets("NEW AC STATE").Select
    Sheets("New AC STATE").Copy Before:=Sheets(1)
    Range("A1:AC34").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
    Application.CutCopyMode = False
    Sheets("NEW AC STATE (2)").Name = "AC STATE"
    Sheets("AC STATE").Move
    Columns("AD:BF").Select
    Selection.Delete Shift:=xlToLeft
    Rows("36:77").Select
    Selection.Delete Shift:=xlUp
    filename = folder & Format(n, "dd-mm-yy") & ".xls"
    ActiveWorkbook.SaveAs filename:=folder & Format(n, "dd-mm-yy") & ".xls", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
    Sheets("NEW AC STATE").Select
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollRow = 1
    Application.CutCopyMode = False
    
'
'
'   To Create BackUp of AC State
'
'
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Dir(ThisWorkbook.Path & "\Back_Up\", vbDirectory) = "" Then
            MkDir (ThisWorkbook.Path & "\Back_Up\")
    End If
    fso.CopyFile ThisWorkbook.Path & "\AC_State.xls", ThisWorkbook.Path & "\Back_Up\AC_State.xls"
    
End Sub

