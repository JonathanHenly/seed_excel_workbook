Attribute VB_Name = "Module1"
Sub Group12_Click()
    dfNewEntry.Show
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
