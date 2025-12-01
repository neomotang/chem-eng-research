Attribute VB_Name = "f9ClearHome"

Sub ClearHome()

'compiled August 2024

'This sub clears the "Home" tab.

Sheet1.Activate
Range("F8", Range("F10")).Value = ""
Range("A21", Range("AI21").End(xlDown)).Select
Selection.Value = ""

Range("A21").Select

End Sub
