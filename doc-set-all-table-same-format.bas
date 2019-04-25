Sub ApplyTableStyle()
Dim t As Table
For Each t In ActiveDocument.Tables
t.Style = "Light Shading - Accent 3" 'Specify table style name here
Next 
End Sub