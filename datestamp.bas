Sub DateStamp()


'Unprotect worksheet for editing
Worksheets("Customer Invoice").Unprotect Password:="Expl0r!ng"

'Insert current date in cell E3
Range("E3") = Date

'Adds bold format to cell E3
Range("E3").Font.Italic = True

'Protects worksheet using the password Expl0r!ng
Worksheets("Customer Invoice").Protect Password:="Expl0r!ng"

End Sub