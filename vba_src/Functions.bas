Attribute VB_Name = "Functions"
Sub Group12_Click()
    dfNewEntry.Show
End Sub

'export specified source codes to destination folder
Private Sub GEN_USE_ExportAllModulesFromProject()
    'https://www.ozgrid.com/forum/forum/help-forums/excel-general/60787-export-all-modules-in-current-project
         'reference to extensibility library
    Dim objMyProj As VBProject
    Dim objVBComp As VBComponent
    Dim destFolder As String
    
    destFolder = "E:\garden\repo\seed_excel_workbook\vba_src\"
    Set objMyProj = Application.VBE.ActiveVBProject
    
    Dim ex_srcs_msg As String
    ex_srcs_msg = "Export Path:" & vbNewLine & " - " & destFolder _
                  & vbNewLine & vbNewLine & "Exported Sources:"
    
    Dim ex_cur As Boolean: ex_cur = False
    For Each objVBComp In objMyProj.VBComponents
        'only export certain sources, not every sheet's source
        Select Case CStr(objVBComp.name)
            Case "ThisWorkbook": ex_cur = True
            Case "dfNewEntry": ex_cur = True
            Case "Functions": ex_cur = True
        End Select
        
        If ex_cur Then
            Dim ex_name As String
            ex_name = objVBComp.name & ".bas"
            
            objVBComp.Export destFolder & ex_name
            
            ex_srcs_msg = ex_srcs_msg & vbNewLine & " - " & ex_name
            
            ex_cur = False
        End If
    Next
    
    MsgBox Prompt:=ex_srcs_msg, title:="Source Exporting Complete"
End Sub

'returns whether a passed in string exists in a range
Function StringMatchedInRange(what As String, rng As Range)
    Dim matched_row As Long
    Dim found_match As Boolean
    found_match = True
    
    On Error GoTo MatchNotFound:
    With Application.WorksheetFunction
        matched_row = .Match(Trim(tbName), rngNames)
    End With
    
ReturnFromFunction:
    StringMatchedInRange = found_match
    Exit Function
    
MatchNotFound:
    found_match = False
    Resume ReturnFromFunction:
End Function
