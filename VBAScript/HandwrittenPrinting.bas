Sub randomFonts()
'
' randomFonts 
    Dim types(0 To 4) As String
    types(0) = "Merkucio Font4You"
    types(1) = "Eskal Font4You"
    types(2) = "Lorenco - Font4You"
    
    For counter = 1 To 1000
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Name = types(Int(Rnd() * 3))
    Selection.Font.Size = Int(Rnd() * 2) + 14
    Selection.Font.Spacing = Rnd() * 0.5 + 0.1
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    
    Select Case (Int(Rnd() * 3))
    Case 1
    Selection.ParagraphFormat.Space1
    Case 2
    Selection.ParagraphFormat.Space15
    Case 3
    Selection.ParagraphFormat.Space2
    End Select
    Next counter
'

End Sub

