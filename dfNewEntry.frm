VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dfNewEntry 
   Caption         =   "Create New Entry"
   ClientHeight    =   5866
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

Dim new_type_ws_was_created As Boolean
Dim DOUBLE_QUOTE As String


Private Sub UserForm_Initialize()
    ThisWorkbook.InitGlobals
    
    'initialize dfNewEntry globals
    'signal that a new type worksheet has not been created yet
    new_type_ws_was_created = False
    DOUBLE_QUOTE = Chr(34)
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws_names() As String: ws_names = GetTypesArray(wb) 'get string array of worksheet names
    
    'initialize type combo box list
    cbType.List = ws_names
    cbType.ListRows = 10 'ArrayLen(ws_names)
    
    ShowNoTypeSelectedLayout
End Sub

Private Sub btnCancel_Click()
    If new_type_ws_was_created Then
        
    End If
    
    'show the packet info sheet if cancel is pressed
    ThisWorkbook.PacketInfoWS.Activate
    
    'unload/exit this form
    Unload dfNewEntry
End Sub

'add VBA worksheet change code to passed in worksheet
Private Function AddWorksheetChangeCode(ws As Worksheet)
    Dim sCode As String
    
    'construct worksheet change string
    sCode = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbNewLine & vbNewLine
    
    sCode = sCode & vbTab & "If Target.Rows.Count = 1 And Target.Columns.Count = Columns.Count Then" & vbNewLine & vbNewLine
    
    sCode = sCode & vbTab & vbTab & "'pass event to workbook" & vbNewLine
    sCode = sCode & vbTab & vbTab & "ThisWorkbook.RowsDeletedFromWorksheet Me, Target" & vbNewLine
    
    sCode = sCode & vbTab & "End If" & vbNewLine & vbNewLine
    
    sCode = sCode & "End Sub" & vbNewLine
    
    'make sure Tools -> References... -> Microsoft Visual Basic for Applications Extensibility 5.3
    'is enabled to insert VBA code into worksheets
    
    'subscript out of range error:
    '-> https://stackoverflow.com/questions/6138689/run-time-error-9-subscript-out-of-range-only-when-excel-vbe-is-closed
    
    'add code to passed in worksheet
    ThisWorkbook.VBProject.VBComponents(ws.CodeName).CodeModule.AddFromString sCode
End Function

Private Sub btnCreate_Click()
    Dim ws As Worksheet
    
    If cbType.Value = "new type" Then 'new type selected
    
        If ValidateNewType Then
            Dim strname As String
            strname = LCase(Trim(tbName.Value))
            
            Application.EnableEvents = False
            
            Set ws = ThisWorkbook.CreateNewWorksheetAndMirrorMaster(strname)
            AddWorksheetChangeCode ws
            
            new_type_ws_was_created = True
            
            'add new type to cbType.List
            cbType.List = GetTypesArray(ThisWorkbook)
            cbType.Value = cbType.List(ArrayLen(cbType.List) - 1)
            
            tbName.Value = ""
            
            Application.EnableEvents = True
        End If
        
    Else 'existing type selected
        Set ws = ThisWorkbook.Sheets(cbType.Value)
        
        If ValidateExistingType(ws) Then
            InsertCreatedEntry ws
            
            new_type_ws_was_created = False
            
            Unload dfNewEntry
            ThisWorkbook.PacketInfoWS.Activate
        End If
    End If
    
    If Not new_type_ws_was_created Then
        ThisWorkbook.UpdateMasterAfterInsert ThisWorkbook.PacketInfoWS, ws
    End If
End Sub

Private Function ValidateExistingType(ws As Worksheet) As Boolean
    Dim v As Boolean: v = True
    
    'check if tbName is empty
    If Trim(tbName.Value) = "" Then
        v = False
        ErrorTextBox tbName
        MsgBox Prompt:="The new type name cannot be empty.", title:="Error"
    End If
    
    If v = False Then
        ValidateExistingType = v
        Exit Function
    End If
    
    ' try to retrieve the product by ID
    Dim rngNames As Range, rngName As Range
    Set rngNames = ws.Range([A4], [A4].End(xlDown))

    Set rngName = rngNames.Find(Trim(tbName), LookIn:=xlValues)
    If rngName Is Nothing Then
        v = True
    Else
        'name already exists, notify and error tbName
        MsgBox "The name, " & Trim(tbName) & " already exists.", title:="Error"
        ErrorTextBox tbName
        v = False
    End If
    
    ValidateExistingType = v
End Function

Private Sub cbType_Change()
    If cbType.Value = "new type" Then
        ShowNoTypeSelectedLayout
        ShowCreateNewTypeLayout
    ElseIf cbType.Value = "" Then
        ShowNoTypeSelectedLayout
    Else
        ShowExistingTypeLayout
        ThisWorkbook.Sheets(cbType.Value).Activate
    End If
End Sub

'gets the next empty cell in a passed in column
Private Function GetNextEmptyCell(ws As Worksheet) As Range
    Dim rng As Range
    Set rng = ws.Columns(1).Find("*", ws.[A1], xlFormulas, , xlByColumns, xlPrevious)
    
    If rng Is Nothing Then
        Set rng = ws.[A3]
    End If
    
    Set GetNextEmptyCell = rng.Offset(1, 0)
End Function

'get data ready and append it to sheet
Private Function InsertCreatedEntry(ByRef ws As Worksheet)
    Dim RefRange As Range
    Set RefRange = GetNextEmptyCell(ws)
    
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
    
    'insert suggestion value
    sug.Value = Trim(tbSuggestions.Value)
    
    'notify user of successful insertion
    MsgBox Trim(tbName.Value) & " was created successfully."
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


Private Sub tbName_Enter()
    ResetTextBox tbName
End Sub

Private Function NewTypeSelected()
    
End Function

Private Function ClearAllButType()
    tbDepth.Value = ""
    tbName.Value = ""
    tbGerm.Value = ""
    tbMaturity.Value = ""
    tbPlant.Value = ""
    tbRow.Value = ""
    tbStart.Value = ""
    tbSuggestions.Value = ""
    cbFull.Value = False
    cbPart.Value = False
End Function


Private Function ShowNoTypeSelectedLayout()
    ClearAllButType
    DisableTextBox tbName
    
    frPacketInfo.Visible = False
    
    btnCreate.Enabled = False
    btnCreate.Top = 60
    
    btnCancel.Top = 60
    
    Me.Height = (318.15 - 264) + btnCancel.Top
    
    cbType.SetFocus
End Function

Private Function ShowCreateNewTypeLayout()
    EnableTextBox tbName
    btnCreate.Enabled = True
    btnCreate.Caption = "Create New Type"
End Function

Private Function ValidateNewType() As Boolean
    Dim v As Boolean: v = True
    
    'check if tbName is empty
    If Trim(tbName.Value) = "" Then
        v = False
        ErrorTextBox tbName
        MsgBox Prompt:="The new type name cannot be empty.", title:="Error"
    End If
    
    If v = False Then
        ValidateNewType = False
        Exit Function
    End If
    
    Dim ws_names() As String
    'get string array of worksheet names
    ws_names = GetTypesArray(ThisWorkbook)
    
    'check if tbName value is already a type
    For Each s In ws_names
        If LCase(Trim(tbName.Value)) = LCase(Trim(s)) Then
            v = False
            ErrorTextBox tbName
            MsgBox Prompt:="The type, " & s & ", already exists.", title:="Error"
        End If
        
        If v = False Then
            Exit For
        End If
    Next s
    
    ValidateNewType = v
End Function

Private Function ShowExistingTypeLayout()
    EnableTextBox tbName
    
    Me.Width = 318.15
    Me.Height = 318.15
    
    btnCreate.Top = 264
    btnCancel.Top = 264
    
    btnCreate.Width = 72.05
    btnCreate.Caption = "Create"
    btnCreate.Enabled = True

    frPacketInfo.Visible = True
End Function

Private Function EnableTextBox(tb As MSForms.TextBox)
    tb.Enabled = True
    tb.BackColor = vbWhite
End Function

Private Function DisableTextBox(tb As MSForms.TextBox)
    tb.Enabled = False
    tb.BackColor = &H8000000F
End Function

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

Public Function ArrayLen(arr As Variant) As Long
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function
