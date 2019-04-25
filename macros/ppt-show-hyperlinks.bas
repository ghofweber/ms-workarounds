Sub ShowMeTheHyperlinks()
' Lists the slide number, shape name and address
' of each hyperlink

    Dim oSl As Slide
    Dim oHl As Hyperlink

    For Each oSl In ActivePresentation.Slides
        For Each oHl In oSl.Hyperlinks
            If oHl.Type = msoHyperlinkShape Then
                MsgBox "HYPERLINK IN SHAPE" _
                    & vbCrLf _
                    & "Slide: " & vbTab & oSl.SlideIndex _
                    & vbCrLf _
                    & "Shape: " & oHl.Parent.Parent.Name _
                    & vbCrLf _
                    & "Address:" & vbTab & oHl.Address _
                    & vbCrLf _
                    & "SubAddress:" & vbTab & oHl.SubAddress
            Else
                ' it's text
                MsgBox "HYPERLINK IN TEXT" _
                    & vbCrLf _
                    & "Slide: " & vbTab & oSl.SlideIndex _
                    & vbCrLf _
                    & "Shape: " & oHl.Parent.Parent.Parent.Parent.Name _
                    & vbCrLf _
                    & "Address:" & vbTab & oHl.Address _
                    & vbCrLf _
                    & "SubAddress:" & vbTab & oHl.SubAddress
            End If
        Next    ' hyperlink
    Next    ' Slide

End Sub