Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    tempColor = Target.Interior.Color
    
    If Target.row <= 30 And Target.Column <= 30 Then
        If tempColor = vbBlack Then Target.Interior.Color = vbWhite
        If tempColor = vbWhite Then Target.Interior.Color = vbBlack
    End If
End Sub
