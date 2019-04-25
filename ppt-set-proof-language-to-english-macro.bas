Public Sub changeLanguage()
    On Error Resume Next
    Dim gi As GroupShapes '<-this was added. used below
    lang = "English"
    'lang = "Norwegian"
    'Determine language selected
    If lang = "English" Then
        lang = msoLanguageIDEnglishUK
    ElseIf lang = "Norwegian" Then
        lang = msoLanguageIDNorwegianBokmol
    End If
    'Set default language in application
    ActivePresentation.DefaultLanguageID = lang

    'Set language in each textbox in each slide
    For Each oSlide In ActivePresentation.Slides
        Dim oShape As Shape
        For Each oShape In oSlide.Shapes
            'Check first if it is a table
            If oShape.HasTable Then
                For r = 1 To oShape.Table.Rows.Count
                    For c = 1 To oShape.Table.Columns.Count
                    oShape.Table.Cell(r, c).Shape.TextFrame.TextRange.LanguageID = lang
                    Next
                Next
            Else
                Set gi = oShape.GroupItems
                'Check if it is a group of shapes
                If Not gi Is Nothing Then
                    If oShape.GroupItems.Count > 0 Then
                        For i = 0 To oShape.GroupItems.Count - 1
                            oShape.GroupItems(i).TextFrame.TextRange.LanguageID = lang
                        Next
                    End If
                'it's none of the above, it's just a simple shape, change the language ID
                Else
                    oShape.TextFrame.TextRange.LanguageID = lang
                End If
            End If
        Next
    Next
End Sub
