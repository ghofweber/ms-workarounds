Option Explicit

Sub FileEmDano()
' Lists the slide number, shape name and address
' of each hyperlink and saves the results to a file:

Dim oSl As Slide
Dim oHl As Hyperlink
Dim sTemp As String
Dim sFileName As String

' Output the results to this file:
sFileName = Environ$("TEMP") & "\" & "HyperlinkList.TXT"

For Each oSl In ActivePresentation.Slides
    For Each oHl In oSl.Hyperlinks
        If oHl.Type = msoHyperlinkShape Then
            sTemp = sTemp & "HYPERLINK IN SHAPE on Slide:" & vbTab & oSl.SlideIndex _
                & vbCrLf _
                & "Shape: " & oHl.Parent.Parent.Name _
                & vbCrLf _
                & "Address:" & vbTab & oHl.Address _
                & vbCrLf _
                & "SubAddress:" & vbTab & oHl.SubAddress & vbCrLf & vbCrLf
        Else
            ' it's text
            sTemp = sTemp & "HYPERLINK IN TEXT on Slide:" & vbTab & oSl.SlideIndex _
                & vbCrLf _
                & "Shape: " & oHl.Parent.Parent.Parent.Parent.Name _
                & vbCrLf _
                & "Address:" & vbTab & oHl.Address _
                & vbCrLf _
                & "SubAddress:" & vbTab & oHl.SubAddress & vbCrLf & vbCrLf
        End If
    Next    ' hyperlink

Next    ' Slide

Call WriteStringToFile(sFileName, sTemp)
Call LaunchFileInNotePad(sFileName)

End Sub

Sub WriteStringToFile(pFileName As String, pString As String)
' Saves the contents of the string pSTring to the file pFileName
    Dim intFileNum As Integer
    intFileNum = FreeFile
                ' change Output to Append if you want to add to an existing file
                ' rather than creating a new file each time
    Open pFileName For Output As intFileNum
    Print #intFileNum, pString
    Close intFileNum

End Sub

Sub LaunchFileInNotePad(pFileName As String)
    Dim lngReturn As Long
    lngReturn = Shell("NOTEPAD.EXE " & pFileName, vbNormalFocus)

End Sub