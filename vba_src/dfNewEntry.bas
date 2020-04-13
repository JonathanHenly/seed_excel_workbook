VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dfNewEntry 
   Caption         =   "Create New Entry"
   ClientHeight    =   6265
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   6181
   OleObjectBlob   =   "dfNewEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dfNewEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'globals
Public newly_created_type_ws As Worksheet
Public new_type_ws_was_created As Boolean
Public inserting_new_entry As Boolean

'constants
Dim DOUBLE_QUOTE As String
Dim FORM_WIDTH As Long
Dim FORM_HEIGHT_EXIST_TYPE As Long
Dim FORM_HEIGHT_TOMATO
Dim BUTTON_TOP_NO_TYPE As Long
Dim BUTTON_TOP_EXIST_TYPE As Long
Dim BUTTON_TOP_TOMATO As Long
Dim FORM_HEIGHT_NO_TYPE As Long
Dim BUTTON_CREATE_WIDTH As Long
Dim BUTTON_CREATE_TEXT As String
Dim BUTTON_CREATE_TEXT_ALT As String
Dim PACKET_INFO_TOP As Long
Dim PACKET_INFO_TOMATO_TOP As Long

Dim LAST_COLUMN As Long

'initializes form, layout, etc. constants
Private Function InitConstants()
    DOUBLE_QUOTE = Chr(34)
    
    FORM_WIDTH = 318.15
    FORM_HEIGHT_EXIST_TYPE = 318.15
    FORM_HEIGHT_TOMATO = FORM_HEIGHT_EXIST_TYPE + 20
    BUTTON_TOP_NO_TYPE = 60
    BUTTON_TOP_EXIST_TYPE = 264
    BUTTON_TOP_TOMATO = BUTTON_TOP_EXIST_TYPE + 20
    FORM_HEIGHT_NO_TYPE = (FORM_WIDTH - BUTTON_TOP_EXIST_TYPE) + BUTTON_TOP_NO_TYPE
    
    BUTTON_CREATE_WIDTH = 72.05
    BUTTON_CREATE_TEXT = "Create"
    BUTTON_CREATE_TEXT_ALT = "Create New Type"
    
    PACKET_INFO_TOP = 60
    PACKET_INFO_TOMATO_TOP = PACKET_INFO_TOP + 20
    
    'last column in worksheets is currently suggestions, col. # 10
    LAST_COLUMN = 10
End Function

'executes before this form is shown, initializes globals, types array and the form layout
Private Sub UserForm_Initialize()
    ThisWorkbook.InitGlobals
    InitConstants
    
    'initialize dfNewEntry globals
    'signals that a new type worksheet has been created
    new_type_ws_was_created = False
    'stores the newly created type worksheet until after its first entry has been inserted in master
    Set newly_created_type_ws = Nothing
    'signals that a new entry is being inserted
    inserting_new_entry = False
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws_names() As String: ws_names = GetTypesArray(wb) 'get string array of worksheet names
    
    'initialize type combo box list
    cbType.List = ws_names
    cbType.ListRows = 10 'ArrayLen(ws_names)
    
    ShowNoTypeSelectedLayout
End Sub

'handles closing of user form via top right corner (x) button
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        CheckForAndDeleteNewlyCreatedTypeWorksheet
        
        'tip: If you want to prevent closing UserForm by Close (×) button in the right-top
        'corner of the UserForm, just uncomment the following line:
        'Cancel = True
    End If
End Sub

'handles cancel button click, returns focus to PacketInfoWS and unloads this form
Private Sub btnCancel_Click()
    CheckForAndDeleteNewlyCreatedTypeWorksheet
    
    'show the packet info sheet if cancel is pressed
    ThisWorkbook.PacketInfoWS.Activate
    
    'unload/exit this form
    Unload dfNewEntry
End Sub

'check for and delete a new type worksheet if first entry was not inserted in master
Private Function CheckForAndDeleteNewlyCreatedTypeWorksheet()
    If Not newly_created_type_ws Is Nothing Then
        Application.DisplayAlerts = False
        newly_created_type_ws.Delete
        Application.DisplayAlerts = True
    End If
End Function

'determinate tomato radio button click event
Private Sub rbDeter_Click()
    ResetRadioButton rbDeter
    ResetRadioButton rbIndeter
    rbDeter.Value = True
End Sub

'indeterminate tomato radio button click event
Private Sub rbIndeter_Click()
    ResetRadioButton rbIndeter
    ResetRadioButton rbDeter
    rbIndeter.Value = True
End Sub

'add VBA worksheet change code to passed in worksheet
Private Function AddWorksheetChangeCode(ws As Worksheet)
    Dim sCode As String
    
    'construct worksheet change event sub string
    sCode = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbNewLine & vbNewLine
    
    'check if user form is inserting a new entry
    sCode = sCode & vbTab & "'check if user form is inserting a new entry" & vbNewLine
    sCode = sCode & vbTab & "If Not dfNewEntry.inserting_new_entry Then" & vbNewLine
    
    'insert callback code for rows being manually deleted from worksheet
    sCode = sCode & vbTab & vbTab & "If Target.Rows.Count >= 1 And Target.Columns.Count = Columns.Count Then" & vbNewLine & vbNewLine
    sCode = sCode & vbTab & vbTab & vbTab & "'pass deleted row index and count to workbook" & vbNewLine
    sCode = sCode & vbTab & vbTab & vbTab & "ThisWorkbook.RowsDeletedFromWorksheet Me, Target" & vbNewLine
    sCode = sCode & vbTab & vbTab & "End If" & vbNewLine
    
    'end new entry insertion check
    sCode = sCode & vbTab & "End If" & vbNewLine & vbNewLine
        
    'end worksheet change event sub
    sCode = sCode & "End Sub" & vbNewLine
    
    'make sure Tools -> References... -> Microsoft Visual Basic for Applications Extensibility 5.3
    'is enabled to insert VBA code into worksheets
    
    'subscript out of range error:
    '-> https://stackoverflow.com/questions/6138689/run-time-error-9-subscript-out-of-range-only-when-excel-vbe-is-closed
    
    'add code to passed in worksheet
    ThisWorkbook.VBProject.VBComponents(ws.CodeName).CodeModule.AddFromString sCode
End Function

'handles new type and new entry creation and calls functions to update master
Private Sub btnCreate_Click()
    Dim ws As Worksheet
    
    If cbType.Value = "new type" Then 'new type selected
        new_type_ws_was_created = False
        
        If ValidateNewType Then
            Dim strname As String
            strname = LCase(Trim(tbName.Value))
            
            Application.EnableEvents = False
            
            Set ws = ThisWorkbook.CreateNewWorksheetAndMirrorMaster(strname)
            AddWorksheetChangeCode ws
            
            'assign new type ws, in case we need to delete it on cancel or close
            Set newly_created_type_ws = ws
            new_type_ws_was_created = True
            
            'add new type to cbType.List
            cbType.List = GetTypesArray(ThisWorkbook)
            cbType.Value = cbType.List(ArrayLen(cbType.List) - 1)
            cbType.Enabled = False
            
            tbName.Value = ""
            
            Application.EnableEvents = True
        End If
        
    Else 'existing type selected
        new_type_ws_was_created = False
        Set ws = ThisWorkbook.Sheets(cbType.Value)
        
        If ValidateExistingType(ws) Then
            'signal new entry insertion
            inserting_new_entry = True
            
            Dim ins_row As Long
            ins_row = InsertCreatedEntry(ws)
            
            'un-signal new entry insertion
            inserting_new_entry = False
            
            'stop showing user form
            Unload dfNewEntry
            
            'set up the copy range for master
            Dim cpy_range As Range
            Set cpy_range = ws.Range(ws.Cells(ins_row, 1), ws.Cells(ins_row, LAST_COLUMN))
            
            ThisWorkbook.UpdateMasterAfterInsert ThisWorkbook.PacketInfoWS, ws, cpy_range
            
            'set newly created type worksheet to nothing after first entry is added to master
            Set newly_created_type_ws = Nothing
            cbType.Enabled = True
                        
            'show master worksheet again
            ThisWorkbook.PacketInfoWS.Activate
        End If
    End If
    
    If Not new_type_ws_was_created Then
    End If
End Sub

'validates create new entry calls, makes sure new entry isn't a duplicate and isn't empty
Private Function ValidateExistingType(ws As Worksheet) As Boolean
    Dim valid As Boolean: valid = True
    Dim name_str As String
    name_str = CStr(Trim(tbName.Value))
    
    'check if tbName is empty
    If name_str = "" Then
        'tbName is empty, notify user
        MsgBox Prompt:="The new type name cannot be empty.", title:="Error"
        'give tbName an error look
        ErrorTextBox tbName
        valid = False
    End If
    
    If valid = False Then
        ValidateExistingType = valid
        Exit Function
    End If
    
    'check if name is already present in type's names
    Dim rngNames As Range
    Set rngNames = ws.Range([A4], [A4].End(xlDown))
    
    If StringMatchedInRange(name_str, rngNames) Then
        'name already exists, notify user
        MsgBox "The name, " & Trim(tbName) & " already exists.", title:="Error"
        'give tbName an error look
        ErrorTextBox tbName
        valid = False
    End If
    
    'if tomato -> validate indeterminate and determinate
    If cbType.Value = "tomato" Then
        If rbIndeter.Value = True And rbDeter.Value = True Then
            MsgBox "Tomatoes cannot be both indeterminate and determinate, please choose one.", _
                title:="Error"
                
            ErrorRadioButton rbIndeter
            ErrorRadioButton rbDeter
            valid = False
        ElseIf rbIndeter.Value = False And rbDeter.Value = False Then
            MsgBox "You must specify if this tomato is indeterminate or determinate.", title:="Error"
            
            ErrorRadioButton rbIndeter
            ErrorRadioButton rbDeter
            valid = False
        End If
    End If
    
    ValidateExistingType = valid
End Function

'handles different types being selected from the type combo box
Private Sub cbType_Change()
    Dim cb_type As String
    cb_type = cbType.Value
    
    If cb_type = "new type" Then
        ShowNoTypeSelectedLayout
        ShowCreateNewTypeLayout
    ElseIf cb_type = "" Then
        ShowNoTypeSelectedLayout
    Else
        ShowExistingTypeLayout
        
        If cb_type = "tomato" Then
            ShowTomatoTypeLayout
        End If
        
        ThisWorkbook.Sheets(cbType.Value).Activate
    End If
End Sub

'gets the next empty cell in a passed in column
Private Function GetNextEmptyCell(ws As Worksheet) As Range
    Dim rng As Range
    Set rng = ws.Columns(1).Find(what:="*", After:=ws.[A1], LookIn:=xlFormulas, _
                                 SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    If rng Is Nothing Then
        Set rng = ws.[A3]
    End If
    
    Set GetNextEmptyCell = rng.Offset(1, 0)
End Function

'validates create new type calls, makes sure new type isn't a duplicate and isn't empty
Private Function ValidateNewType() As Boolean
    Dim valid As Boolean: valid = True
    Dim name_str As String
    
    name_str = CStr(LCase(Trim(tbName.Value)))
    
    'check if tbName is empty
    If name_str = "" Then
        valid = False
        ErrorTextBox tbName
        MsgBox Prompt:="The new type name cannot be empty.", title:="Error"
    End If
    
    If valid = False Then
        ValidateNewType = False
        Exit Function
    End If
    
    Dim ws_names() As String
    'get string array of worksheet names
    ws_names = GetTypesArray(ThisWorkbook)
    
    Dim s As Variant
    'check if tbName value is already a type
    For Each s In ws_names
        If LCase(Trim(tbName.Value)) = LCase(Trim(s)) Then
            valid = False
            ErrorTextBox tbName
            MsgBox Prompt:="The type, " & s & ", already exists.", title:="Error"
        End If
        
        If valid = False Then
            Exit For
        End If
    Next s
    
    ValidateNewType = valid
End Function

'gets form data ready and inserts it alphabetically, by name, in the passed in sheet
Private Function InsertCreatedEntry(ByRef ws As Worksheet) As Long
    Dim ins_index As Long
    ins_index = FindNameRowInsertIndex(ws, 4, CStr(tbName.Value))
    
    'insert new row above ins_index
    InsertRowAboveIndex ws, ins_index
    
    'reference the newly inserted row
    Dim RefRange As Range
    Set RefRange = ws.Cells(ins_index, 1)
    
    Set n = RefRange 'name
    Set dtg = RefRange.Offset(0, 1) 'days to germination
    Set sd = RefRange.Offset(0, 2) 'seed depth
    Set wts = RefRange.Offset(0, 3) 'when to start
    Set dtm = RefRange.Offset(0, 4) 'days to maturity
    Set sr = RefRange.Offset(0, 5) 'row spacing
    Set sp = RefRange.Offset(0, 6) 'plant spacing
    Set se = RefRange.Offset(0, 7) 'sun exposure
    Set mh = RefRange.Offset(0, 8) 'mature height
    Set sug = RefRange.Offset(0, 9) 'suggestions
    
    'insert the name value
    n.Value = Trim(tbName.Value)
    
    'insert days till germination
    dtg_str = Trim(tbGerm.Value)
    If Not dtg_str = "" Then
        dtg_str = dtg_str & " days"
    End If
    dtg.Value = dtg_str
    
    'insert seed sowing depth
    sd_str = Trim(tbDepth.Value)
    If Not sd_str = "" Then
        sd_str = sd_str & DOUBLE_QUOTE
    End If
    sd.Value = sd_str
    
    'insert when to start weeks
    wts_str = Trim(tbStart.Value)
    If Not dtg_str = "" Then
        wts_str = wts_str & " weeks"
    End If
    wts.Value = wts_str
    
    'insert days till maturity
    dtm_str = Trim(tbMaturity.Value)
    If Not dtm_str = "" Then
        dtm_str = dtm_str & " days"
    End If
    dtm.Value = dtm_str
    
    'insert row spacing value
    sr_str = Trim(tbRow.Value)
    If Not sr_str = "" Then
        sr_str = sr_str & DOUBLE_QUOTE
    End If
    sr.Value = sr_str
    
    'insert plant spacing value
    sp_str = Trim(tbPlant.Value)
    If Not sp_str = "" Then
        sp_str = sp_str & DOUBLE_QUOTE
    End If
    sp.Value = sp_str
    
    'insert full, partial or full/partial sun
    If cbFull.Value = True And cbPart.Value = True Then
        se.Value = "full/part"
    ElseIf cbFull.Value = True Then
        se.Value = "full"
    ElseIf cbPart.Value = True Then
        se.Value = "part"
    Else
        se.Value = "full"
    End If
    
    'insert mature height value
    mh_str = Trim(tbHeight.Value)
    If Not mh_str = "" Then
        mh_str = mh_str & DOUBLE_QUOTE
    End If
    mh.Value = mh_str
    
    Dim indet As String
    indet = ""
    If cbType.Value = "tomato" Then
        If rbIndeter.Value = True Then
            indet = "Indeterminate. "
        Else
            indet = "Determinate. "
        End If
    End If
    'insert suggestion value, if tomato then prepend in/determinate
    sug.Value = Trim(indet & tbSuggestions.Value)
    
    'notify user of successful insertion
    MsgBox Trim(tbName.Value) & " was created successfully."
    
    InsertCreatedEntry = ins_index
End Function

Private Function FindNameRowInsertIndex(ws As Worksheet, start As Long, name_str As String) As Long
    Dim index_found As Boolean
    Dim index As Long
    Dim cur_name_str As String
    Dim res As Long
    
    index_found = False
    index = start
    Do While Not index_found
        cur_name_str = CStr(ws.Cells(index, 1).Value)
        
        If cur_name_str <> "" Then
            res = StrComp(name_str, cur_name_str)
            
            If res = -1 Then 'name_str < cur_name_str alphabetically
                index_found = True
                
            ElseIf res = 1 Then 'name_str > cur_name_str alphabetically
                'iterate index
                index = index + 1
            ElseIf res = 0 Then 'name_str = cur_name_str
                'we should not reach this point, due to no duplicate names
                index_found = True
            End If
            
        ElseIf cur_name_str = "" Then
            'reached end of names column
            index_found = True
        End If
        
    Loop
    
    FindNameRowInsertIndex = index
End Function

'insert a new row above the specfied row index
Function InsertRowAboveIndex(ws As Worksheet, row_index As Long)
    
    ws.Rows(row_index).Insert Shift:=xlDown, _
        CopyOrigin:=xlFormatFromRightOrBelow
End Function

'give passed in textbox a red border to signal error
Private Function ErrorTextBox(tb As MSForms.TextBox)
    tb.SpecialEffect = fmSpecialEffectFlat
    tb.BorderColor = &HFF
    tb.BorderStyle = fmBorderStyleSingle
End Function

'return passed in textbox to its default look
Private Function ResetTextBox(tb As MSForms.TextBox)
    tb.BorderColor = &H80000006
    tb.BorderStyle = fmBorderStyleNone
    tb.SpecialEffect = fmSpecialEffectSunken
End Function

'give passed in radio button a red background to signal error
Private Function ErrorRadioButton(rb As MSForms.OptionButton)
    rb.BackColor = &HFF
End Function

'return passed in radio button to its default look
Private Function ResetRadioButton(rb As MSForms.OptionButton)
    rb.BackColor = &H8000000F
    rb.SpecialEffect = fmButtonEffectSunken
End Function

'return name textbox to its default look when entered, resets error textbox
Private Sub tbName_Enter()
    ResetTextBox tbName
End Sub

'clears all form data other than type
Private Function ClearAllButType()
    tbDepth.Value = ""
    tbName.Value = ""
    tbGerm.Value = ""
    tbMaturity.Value = ""
    tbRow.Value = ""
    tbPlant.Value = ""
    tbStart.Value = ""
    tbHeight.Value = ""
    tbSuggestions.Value = ""
    cbFull.Value = False
    cbPart.Value = False
    rbIndeter.Value = 0
    rbDeter.Value = 0
End Function


'type combo box has nothing selected, so show that layout
Private Function ShowNoTypeSelectedLayout()
    ClearAllButType
    DisableTextBox tbName
    
    rbIndeter.Visible = False
    rbDeter.Visible = False
    frPacketInfo.Visible = False
    frPacketInfo.Top = PACKET_INFO_TOP
    
    btnCreate.Enabled = False
    btnCreate.Top = BUTTON_TOP_NO_TYPE
    
    btnCancel.Top = BUTTON_TOP_NO_TYPE
    
    Me.Width = FORM_WIDTH
    Me.Height = FORM_HEIGHT_NO_TYPE
    
    cbType.SetFocus
End Function

'type combo box has "new entry" selected, so show that layout
Private Function ShowCreateNewTypeLayout()
    Me.Width = FORM_WIDTH
    Me.Height = FORM_HEIGHT_NO_TYPE
    
    EnableTextBox tbName
    
    btnCreate.Enabled = True
    btnCreate.Caption = BUTTON_CREATE_TEXT_ALT
    
    rbIndeter.Visible = False
    rbDeter.Visible = False
    
    tbName.SetFocus
End Function

'type combo box has an existing type selected, other than tomato, so show that layout
Private Function ShowExistingTypeLayout()
    EnableTextBox tbName
    
    Me.Width = FORM_WIDTH
    Me.Height = FORM_HEIGHT_EXIST_TYPE
    
    btnCreate.Top = BUTTON_TOP_EXIST_TYPE
    btnCancel.Top = BUTTON_TOP_EXIST_TYPE
    
    btnCreate.Width = BUTTON_CREATE_WIDTH
    btnCreate.Caption = BUTTON_CREATE_TEXT
    btnCreate.Enabled = True
    
    frPacketInfo.Top = PACKET_INFO_TOP
    frPacketInfo.Visible = True
    
    rbIndeter.Visible = False
    rbDeter.Visible = False
End Function

'type combo box has tomato selected, so show that layout
Private Function ShowTomatoTypeLayout()
    Me.Height = FORM_HEIGHT_TOMATO
    
    btnCreate.Top = BUTTON_TOP_TOMATO
    btnCancel.Top = BUTTON_TOP_TOMATO
    
    frPacketInfo.Top = PACKET_INFO_TOMATO_TOP
    
    rbIndeter.Visible = True
    rbDeter.Visible = True
    
End Function

'puts a passed in textbox to an enabled state with a white background
Private Function EnableTextBox(tb As MSForms.TextBox)
    tb.Enabled = True
    tb.BackColor = vbWhite
End Function

'puts a passed in textbox to a disabled state with a gray background
Private Function DisableTextBox(tb As MSForms.TextBox)
    tb.Enabled = False
    tb.BackColor = &H8000000F
End Function

'returns an array of strings populated with all the names of the worksheets, other than master
Private Function GetTypesArray(ByRef wb As Workbook) As String()
    Dim ws_names() As String
    '-1 due to wb.PacketInfoWS
    ReDim ws_names(wb.Worksheets.count - 1)
    Dim ws As Worksheet
    
    ws_names(0) = "new type"
    Dim ncount As Long: ncount = 1
    Dim x As Long
    For x = 1 To wb.Worksheets.count
        Set ws = wb.Worksheets(x)
        If Not ws.name = wb.PacketInfoWS.name Then
            ws_names(ncount) = ws.name
            ncount = ncount + 1
        End If
    Next x
    
    GetTypesArray = ws_names
End Function

'returns the length of a passed in array
Public Function ArrayLen(arr As Variant) As Long
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function
