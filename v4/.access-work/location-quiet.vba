Private Sub LoadLocation(ByVal selectedID As Variant, Optional ByVal legacyTag As Variant = "")
    HideLocationMatches
    locationID = selectedID
    locationLabel = Null
    If Not IsNull(selectedID) Then
        locationLabel = DLookup("LocationTag", "LocationTags", "LocationTagID=" & CLng(selectedID))
    ElseIf Len(Nz(legacyTag, "")) > 0 And editedID <> 0 Then
        locationID = 0
        locationLabel = legacyTag
    End If
    Me.txtLocation = locationLabel
End Sub

Private Sub HideLocationMatches()
    Me.lstLocationMatches.Visible = False
    If editedID = 0 Or Me.txtLocation.Locked Then
        Me.TimerInterval = 0
    Else
        Me.TimerInterval = 1000
    End If
End Sub

Private Sub txtLocation_Change()
    On Error GoTo Failed
    Dim sql As String
    sql = LocationSearchRows(Me.txtLocation.Text)
    If sql <> Me.lstLocationMatches.RowSource Then Me.lstLocationMatches.RowSource = sql
    If Me.lstLocationMatches.ListCount = 0 Then
        HideLocationMatches
    Else
        Me.lstLocationMatches.Visible = True
        Me.TimerInterval = 200
    End If
    Exit Sub
Failed:
    Me.lblStatus.Caption = "Location search unavailable: " & Err.Description
End Sub

Private Sub txtLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not Me.lstLocationMatches.Visible Then Exit Sub
    Select Case KeyCode
        Case vbKeyDown
            KeyCode = 0
            Me.lstLocationMatches.SetFocus
            Me.lstLocationMatches = Me.lstLocationMatches.ItemData(0)
        Case vbKeyEscape
            KeyCode = 0
            HideLocationMatches
    End Select
End Sub

Private Sub lstLocationMatches_Click()
    ChooseLocation
End Sub

Private Sub lstLocationMatches_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            KeyCode = 0
            ChooseLocation
        Case vbKeyEscape
            KeyCode = 0
            Me.txtLocation.SetFocus
            HideLocationMatches
    End Select
End Sub

Private Sub ChooseLocation()
    If IsNull(Me.lstLocationMatches) Then Exit Sub
    locationID = Me.lstLocationMatches.Value
    locationLabel = Me.lstLocationMatches.Column(1)
    Me.txtLocation = locationLabel
    Me.txtLocation.SetFocus
    Me.txtLocation.SelStart = Len(locationLabel)
    HideLocationMatches
    MarkChanged
End Sub

Private Function ChosenLocationID() As Variant
    ChosenLocationID = Null
    If StrComp(Nz(Me.txtLocation.Value, ""), Nz(locationLabel, ""), vbBinaryCompare) = 0 Then ChosenLocationID = locationID
End Function
