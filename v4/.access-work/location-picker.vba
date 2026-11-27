Private Function LocationRows(Optional ByVal selectedID As Long = -1) As String
    LocationRows = "SELECT LocationTagID, LocationTag, ObjectType, LongTag FROM LocationTags WHERE LocationTagID=" & selectedID
End Function

Private Sub LoadLocation(ByVal selectedID As Variant, Optional ByVal legacyTag As Variant)
    Me.txtLocation = Null
    If Not IsNull(selectedID) Then
        Me.txtLocation.RowSource = LocationRows(CLng(selectedID))
        Me.txtLocation = selectedID
    ElseIf Len(Nz(legacyTag, "")) > 0 And editedID <> 0 Then
        Me.txtLocation.RowSource = "SELECT 0 AS TagID, LocationTag, '' AS ObjectType, 'Previous value (unlinked)' AS LongTag FROM DailyReports WHERE ReportID=" & editedID
        Me.txtLocation = 0
    Else
        Me.txtLocation.RowSource = LocationRows()
    End If
End Sub

Public Function LocationSearchRows(ByVal typedTag As String) As String
    Dim key As String, pattern As String, position As Long, character As String
    key = UCase$(Replace(Replace(Replace(Trim$(typedTag), "-", ""), " ", ""), vbTab, ""))
    If Len(key) < 3 Or Len(key) > 100 Then
        LocationSearchRows = LocationRows()
        Exit Function
    End If
    For position = 1 To Len(key)
        character = Mid$(key, position, 1)
        Select Case character
            Case "'": pattern = pattern & "''"
            Case "[", "*", "?", "#", "%", "_": pattern = pattern & "[" & character & "]"
            Case Else: pattern = pattern & character
        End Select
    Next
    LocationSearchRows = "SELECT TOP 50 LocationTagID, LocationTag, ObjectType, LongTag FROM LocationTags WHERE SearchKey LIKE '" & pattern & "*' ORDER BY SearchKey, LocationTagID"
End Function

Private Sub txtLocation_Change()
    If updatingLocation Or Me.txtLocation.Locked Then Exit Sub
    On Error GoTo Failed
    Dim typedTag As String, caret As Long, sql As String
    typedTag = Me.txtLocation.Text
    caret = Me.txtLocation.SelStart
    sql = LocationSearchRows(typedTag)
    If Me.txtLocation.RowSource = sql Then Exit Sub
    updatingLocation = True
    Me.txtLocation.RowSource = sql
    Me.txtLocation.Text = typedTag
    Me.txtLocation.SelStart = caret
    If Me.txtLocation.ListCount > 0 Then Me.txtLocation.Dropdown
    updatingLocation = False
    Exit Sub
Failed:
    updatingLocation = False
    Me.lblStatus.Caption = "Location search unavailable: " & Err.Description
End Sub

Private Sub txtLocation_NotInList(NewData As String, Response As Integer)
    Response = acDataErrContinue
    MsgBox "Type at least 3 characters and select a location tag from the list.", vbExclamation, "Location tag"
End Sub
