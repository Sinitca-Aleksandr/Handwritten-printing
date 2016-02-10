Sub randomFonts()
'
' randomFonts 
    Dim types(0 To 2) As String
    types(0) = "Merkucio Font4You"
    types(1) = "Eskal Font4You"
    types(2) = "Lorenco - Font4You"
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = "  "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0
    
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

