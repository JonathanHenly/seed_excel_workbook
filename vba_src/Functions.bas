Attribute VB_Name = "Functions"

'BEGIN MODULE
Option Explicit

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
    
    destFolder = ActiveWorkbook.Path & "\vba_src\"
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
        matched_row = .Match(Trim(tbName), rng)
    End With
    
ReturnFromFunction:
    StringMatchedInRange = found_match
    Exit Function
    
MatchNotFound:
    found_match = False
    Resume ReturnFromFunction:
End Function

'returns True if the exact passed in string exists in a row of contigous non-empty cells,
'otherwise False
Function StringExistsInCol(ws As Worksheet, str As String, roff As Long, coff As Long) _
    As Boolean
    
    StringExistsInCol = (Not FindStringInCol(ws, str, roff, coff) Is Nothing)
End Function

'returns the range of a string found in a row of contigous non-empty cells, or Nothing
'if the exact string is not found
Function FindStringInCol(ws As Worksheet, str As String, roff As Long, coff As Long) _
    As Range
    
    Dim cur_row As Long
    Dim cur_cell_str As String
    Dim res As Long
    Dim found_range As Range
    Dim cur_cell As Range
    
    Set found_range = Nothing
    cur_row = roff
    Set cur_cell = ws.Cells(cur_row, coff)
    
    Do While found_range Is Nothing
        Set cur_cell = ws.Cells(cur_row, coff)
        cur_cell_str = CStr(cur_cell.Value)
        
        If cur_cell_str <> "" Then
            If str <> "" Then
                res = StrComp(str, cur_cell_str)
                
                If res = 0 Then
                    'we found an exact match
                    Set found_range = cur_cell
                Else
                    'move to next row
                    cur_row = cur_row + 1
                End If
            Else
                'special case for passed in empty string
                cur_row = cur_row + 1
            End If
        ElseIf cur_cell_str = "" Then
            'special case for passed in empty string
            If str = "" Then
                Set found_range = cur_cell
            End If
            
            'reached an empty cell, aka. the end
            Exit Do
        End If
        
    Loop
    
    Set FindStringInCol = found_range
End Function

'returns True if the exact passed in string exists in a row, skipping over empty merged
'cells, otherwise False
Function StringExistsInColWithMergedRows(ws As Worksheet, str As String, roff As Long, coff As Long) _
    As Boolean
    
    StringExistsInColWithMergedRows = (Not FindStringInColWithMergedRows(ws, str, roff, coff) Is Nothing)
End Function

'returns the range of a string found in a row, skipping over empty merged cells,
'or Nothing if the exact string is not found
Function FindStringInColWithMergedRows(ws As Worksheet, str As String, roff As Long, _
    coff As Long) As Range
    
    Dim cur_row As Long
    Dim cur_cell_str As String
    Dim res As Long
    Dim found_range As Range
    Dim cur_cell As Range
    
    Set found_range = Nothing
    cur_row = roff
    Set cur_cell = ws.Cells(cur_row, coff)
    
    Do While found_range Is Nothing
        Set cur_cell = ws.Cells(cur_row, coff)
        cur_cell_str = CStr(cur_cell.Value)
        
        If cur_cell_str <> "" Then
            If str <> "" Then
                res = StrComp(str, cur_cell_str)
                
                If res = 0 Then
                    'we found an exact match
                    Set found_range = cur_cell
                Else
                    'move to next row after merged rows
                    cur_row = cur_row + cur_cell.MergeArea.Rows.count
                End If
            Else
                'special case for passed in empty string
                cur_row = cur_row + cur_cell.MergeArea.Rows.count
            End If
        ElseIf cur_cell_str = "" Then
            'special case for passed in empty string
            If str = "" Then
                Set found_range = cur_cell
            End If
            
            'reached an empty cell after merged rows, aka. the end
            Exit Do
        End If
        
    Loop
    
    Set FindStringInColWithMergedRows = found_range
End Function
