VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Public GLOBALS_SET As Boolean
Public TodayDate As String
Public PacketInfoWS As Worksheet


'initialize global variables
Private Sub Workbook_Open()
    InitGlobals
End Sub

Function InitGlobals()
    'exit InitGlobals if globals are already set
    If GLOBALS_SET Then
        Exit Function
    Else
        GLOBALS_SET = True
    End If
    
    TodayDate = Date
    TodayDate = WorksheetFunction.Substitute(TodayDate, "/", "-")
    
    Dim p_info As String: p_info = "master"
    
    If PacketInfoWS Is Nothing Then
        Set PacketInfoWS = ThisWorkbook.Sheets(p_info)
        If PacketInfoWS Is Nothing Then
            MsgBox "No worksheet named " & p_info & " was found."
            Exit Function 'possible way of handling no worksheet found
        End If
    End If
    
End Function

Function RowsDeletedFromWorksheet(ws As Worksheet, Target As Range)
    MsgBox "Row(s) " & Target.row & " deleted from " & ws.name
    MsgBox Target.Rows.count & " row(s) deleted from " & ws.name
End Function

Function UpdateMasterAfterInsert(ma As Worksheet, ws As Worksheet, cpy_range As Range)
    Dim this_type_str As String
    this_type_str = Trim(ws.[A1].Value)
    
    'get all types listed in master
    Dim all_types As Range
    Set all_types = ma.Range(ma.[A4], ma.[A4].End(xlDown))
    
    'find the address of ws's type
    Dim this_type As Range
    Set this_type = all_types.Find(this_type_str, LookIn:=xlValues)
    
    If this_type Is Nothing Then
        'type doesn't exist in master, need to insert it
        InsertNewTypeEntry ma, this_type_str, cpy_range
    Else
        'type already exists, insert new entry
        InsertExistingTypeEntry ma, this_type, cpy_range
    End If
    
End Function

Private Function InsertNewTypeEntry(ma As Worksheet, type_str As String, cpy_range As Range)
    Dim row_insert_index As Long
    row_insert_index = FindRowInsertIndex(ma, 4, type_str)
    
    Dim row_insert_range As Range
    Set row_insert_range = ma.Cells(row_insert_index, 1)
    
    MsgBox "row_insert_range.Row = " & row_insert_range.row & "  col count: " & row_insert_range.Columns.count
    
    row_insert_range.EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    
    row_insert_range.Offset(-1, 0).Value = type_str
    row_insert_range.Offset(-1, 0).HorizontalAlignment = xlCenter
    row_insert_range.Offset(-1, 0).VerticalAlignment = xlCenter
    
    'insert type's first entry
    cpy_range.copy
    this_type.Offset(-1, 1).PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
    
    InsertTypesFirstEntry row_insert_range, cpy_range
End Function

Public Function FindRowInsertIndex(ws As Worksheet, start As Long, type_str As String) As Long
    If type_str = "" Then
        MsgBox "The new type string cannot be empty", title:="Error"
        Exit Function
    End If
    
    Dim index_found As Boolean
    Dim index As Long
    Dim cur_type_str As String
    Dim mrc As Long
    Dim res As Long
            
    index_found = False
    index = start
    Do While Not index_found
        cur_type_str = CStr(ws.Cells(index, 1).Value)
        
        mrc = ws.Cells(index, 1).MergeArea.Rows.count
        
        If cur_type_str <> "" Then
            res = StrComp(type_str, cur_type_str)
            
            If res = -1 Then 'type_str < cur_type_str alphabetically
                index_found = True
                
            ElseIf res = 1 Then 'type_str > cur_type_str alphabetically
                'iterate index
                index = index + mrc
            ElseIf res = 0 Then 'type_str = cur_type_str
                'we should not reach this point, due to no duplicate types
                index_found = True
            End If
            
        ElseIf cur_type_str = "" Then
            'reached end of types column
            index_found = True
        End If
        
    Loop
    
    FindRowInsertIndex = index
End Function

Private Function InsertExistingTypeEntry(ma As Worksheet, this_type As Range, _
                                         cpy_range As Range)
    MsgBox "This Far!"
    Dim mrc As Long
    mrc = GetMergedRowCount(this_type)
    this_type.UnMerge
    
    'check if this is a type with no entries (i.e. newly created)
    If ma.Cells(this_type.row, 2).Value = "" Then
        InsertTypesFirstEntry ma, this_type, cpy_range
    End If
    
    'DEBUG
    Exit Function
    
    InsertNameAbove ma, this_type, cpy_range
    
    InsertRowBelow this_type, cpy_range
    
    Dim new_range As Range
    Set new_range = ma.Range(this_type.Cells(1, 1), this_type.Cells(mrc + 1, 1))
    new_range.Merge 'Across:=xlCenterAcrossSelection
End Function

Private Function InsertNameAbove(ma As Worksheet, ins_range As Range, cpy_range As Range)
    Dim mrc As Long
    mrc = GetMergedRowCount(ins_range)
    ins_range.UnMerge
    
    Dim ins_row As Long
    
    ins_row = ins_range.row + (mrc - 1)
    
    MsgBox "ins_range.Row = " & ins_range.row & vbNewLine _
         & "cpy_range.Row = " & cpy_range.row & vbNewLine _
         & "mrc = " & mrc & vbNewLine _
         & "ins_row = " & ins_row
    'DEBUG
    Exit Function
    
    
    If (cpy_range.row - 4) > mrc Then
        ins_row = ins_range.row + (mrc - 1)
    Else
        ins_row = ins_range.row + (cpy_range.row - 4)
        
        Dim type_value As String: type_value = ""
        If ins_row = ins_range.row Then
            type_value = CStr(ma.Cells(ins_row, 1).Value)
            ma.Cells(ins_row, 1).Value = ""
            
            ma.Cells(ins_row, 1).EntireRow.Insert Shift:=xlDown, _
                CopyOrigin:=xlFormatFromRightOrBelow
            
            ma.Cells(ins_row, 1).Value = type_value
            
            cpy_range.copy
            
            ma.Cells(ins_row, 2).PasteSpecial Paste:=xlPasteValues
        End If
    End If
    
    
End Function

Function InsertRowAbove(ins_range As Range, cpy_range As Range)
    ins_range.Offset(-1).EntireRow.Insert Shift:=xlDown, _
                                         CopyOrigin:=xlFormatFromRightOrBelow
    
    cpy_range.copy
    
    'ins_range.Offset(1, 1).PasteSpecial xlPasteFormats
    ins_range.Offset(-1, 1).PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
End Function

Function InsertRowBelow(ins_range As Range, cpy_range As Range)
    ins_range.Offset(1).EntireRow.Insert Shift:=xlUp, _
                                         CopyOrigin:=xlFormatFromLeftOrAbove
    
    cpy_range.copy
    
    'ins_range.Offset(1, 1).PasteSpecial xlPasteFormats
    ins_range.Offset(1, 1).PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
End Function

Private Function MergeSheetHeaders(master As Worksheet, ws As Worksheet, cols As Long)
    Dim mtitle As Range
    Dim title As Range
    Set mtitle = master.Range("A1:A1")
    Set title = ws.Range(ws.Cells(1, 1), ws.Cells(1, cols))
    
    title.Merge
    
    Dim count As Long
    Dim roff As Long
    count = 2
    roff = 2
    
    Do While master.Cells(roff, count).Value <> ""
        Dim mr As Range
        Set mr = master.Cells(roff, count)
        
        Dim wr As Range
        Dim rcount As Long
        
        rcount = GetMergedRowCount(mr)
        If rcount > 1 Then
            Set wr = ws.Range( _
                ws.Cells(roff, count - 1), ws.Cells(roff + rcount - 1, count - 1))
            
            wr.Merge
        End If
        
        Dim colcount As Long
        colcount = GetMergedColCount(mr)
        If colcount > 1 Then
            Set wr = ws.Range( _
                ws.Cells(roff, count - 1), ws.Cells(roff, count - 1 + colcount - 1))
            
            wr.Merge
            count = count + colcount - 1
        End If
        
        If rcount = 1 And colcount = 1 Then
            Set wr = ws.Range(ws.Cells(roff, count - 1))
        End If
        
        wr.Value = mr.Value
        
        count = count + 1
    Loop
End Function

Function CreateNewWorksheetAndMirrorMaster(wsname As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add( _
                 After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    
    ws.name = wsname
    
    MirrorMasterHeaders ws
    
    Set CreateNewWorksheetAndMirrorMaster = ws
End Function

Private Function MirrorMasterHeaders(ByRef ws As Worksheet)
    Dim master As Worksheet
    Set master = PacketInfoWS
    
    If master Is Nothing Then
        MsgBox "No worksheet named 'master' was found."
        Exit Function 'possible way of handling no worksheet was set
    End If
    
    Dim roff As Long
    roff = 2
    Dim coff As Long
    coff = 2
    Dim rend As Long
    rend = 3
    
    Application.ScreenUpdating = False
        
    Dim count As Long
    count = FindLastHeaderIndex(master, roff, coff)
    
    Dim matitle As Range
    Set matitle = master.Range("A1", master.Cells(1, count - 1))
    
    Dim maheaders As Range
    Set maheaders = master.Range(master.Cells(roff, coff), master.Cells(rend, count))
    
    Dim wstitle As Range, wsheaders As Range
    
    Set wstitle = ws.Range("A1", ws.Cells(1, count - 1))
    MirrorTitleFormattingFromMaster matitle, wstitle, ws.name
            
    Set wsheaders = ws.Range(ws.Cells(roff, 1), ws.Cells(rend, count))
    MirrorHeadersFromMaster maheaders, wsheaders
            
    MirrorColumnAlignments master, ws, coff, count
    
    Application.ScreenUpdating = True
    
    Application.CutCopyMode = False
End Function

Sub MirrorMasterHeadersAll()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim master As Worksheet
    Set master = PacketInfoWS
    
    If master Is Nothing Then
        MsgBox "No worksheet named 'master' was found."
        Exit Sub 'possible way of handing no worksheet was set
    End If
    
    Dim roff As Long
    roff = 2
    Dim coff As Long
    coff = 2
    Dim rend As Long
    rend = 3
    
    Dim count As Long
    count = FindLastHeaderIndex(master, roff, coff)
    
    Dim matitle As Range
    Set matitle = master.Range("A1", master.Cells(1, count - 1))
    
    Dim maheaders As Range
    Set maheaders = master.Range(master.Cells(roff, coff), master.Cells(rend, count))
    
    Dim ws As Worksheet
    Dim wstitle As Range, wsheaders As Range
    'loop over worksheets excluding PacketInfoWS
    For Each ws In wb.Worksheets
        If Not ws.name = PacketInfoWS.name Then
            ResetWorksheet ws
            Set wstitle = ws.Range("A1", ws.Cells(1, count - 1))
            MirrorTitleFormattingFromMaster matitle, wstitle, ws.name
            
            Set wsheaders = ws.Range(ws.Cells(roff, 1), ws.Cells(rend, count))
            MirrorHeadersFromMaster maheaders, wsheaders
            
            MirrorColumnAlignments master, ws, coff, count
        End If
    Next ws
    
End Sub

Private Function ResetWorksheet(ByRef ws As Worksheet)
    Dim SaveCalcState
    SaveCalcState = Application.Calculation
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    ws.UsedRange.Delete
    Application.ScreenUpdating = True
    Application.Calculation = SaveCalcState
End Function

Private Function MirrorTitleFormattingFromMaster(ByRef matitle As Range, _
                                                 ByRef wstitle As Range, _
                                                 ByRef name As String)
    'copy/paste title from master
    matitle.copy
    wstitle.PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    
    'adjust title row height
    wstitle.RowHeight = matitle.RowHeight
    
    'set title of ws to ws.Name proper
    wstitle.Value = Application.Proper(name)
End Function

Private Function MirrorHeadersFromMaster(ByRef maheaders As Range, _
                                         ByRef wsheaders As Range)
    'copy/paste headers from master
    maheaders.copy
    
    wsheaders.PasteSpecial Paste:=xlPasteColumnWidths
    wsheaders.PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    
    'adjust header row heights
    wsheaders.RowHeight = maheaders.RowHeight
End Function

Private Function MirrorColumnAlignments(ByRef ma As Worksheet, _
                                        ByRef ws As Worksheet, _
                                        coff As Long, cend As Long)
    Dim i As Long
    For i = coff To cend
        Dim wscln As String
        wscln = ColumnLetter(i - (coff - 1))
        Dim macln As String
        macln = ColumnLetter(i)
        
        Dim wsr As Range: Set wsr = ws.Range(wscln & "4:" & wscln & "100")
        Dim mar As Range: Set mar = ma.Range(macln & "4:" & macln & "4")
        wsr.HorizontalAlignment = mar.HorizontalAlignment
        wsr.VerticalAlignment = mar.VerticalAlignment
    Next i
End Function

'returns the letter of the passed in column index
Private Function ColumnLetter(ByVal ColumnNumber As Long) As String
    Dim n As Long
    Dim c As Byte
    Dim s As String

    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    
    ColumnLetter = s
End Function

Private Function FindLastHeaderIndex(ByRef ws As Worksheet, _
                                     roff As Long, coff As Long) As Long
    Dim cols As Long
    
    cols = coff
    Dim colcount As Long
    
    Do While ws.Cells(roff, cols).Value <> ""
        'colcount = ws.Cells(roff, cols).MergeArea.Columns.count
        colcount = GetMergedColCount(ws.Cells(roff, cols))
        If colcount > 1 Then
            cols = cols + colcount - 1
        End If
        
        cols = cols + 1
    Loop
    
    FindLastHeaderIndex = cols - coff + 1
End Function

'gets the next empty row, with an offset, in a passed in column
Function NextEmptyRowByCol(ws As Worksheet, row_off As Long, col As Long) As Long
    Dim rng As Range
    Set rng = ws.Columns(col).Find(what:="*", After:=ws.Cells(row_off, col), _
        LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    Dim row As Long
    If rng Is Nothing Then
        row = 4
    Else
        row = rng.row
    End If
    
    NextEmptyRowByCol = row
End Function

Private Function GetMergedColCount(r As Range) As Long
    GetMergedColCount = r.Cells.MergeArea.Columns.count
End Function

Private Function GetMergedRowCount(r As Range) As Long
    GetMergedRowCount = r.Cells.MergeArea.Rows.count
End Function
