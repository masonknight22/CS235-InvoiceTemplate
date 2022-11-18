Sub ClearInvoice()

'Unprotect worksheet for editing.
Worksheets("Customer Invoice").Unprotect Password:="Expl0r!ng"

'
' ClearInvoice Macro
' This macro clears existing values in the current invoice.
'

'
    Range("C13:E13").Select
    Selection.ClearContents
    Range("C7:C11").Select
    Selection.ClearContents
    Range("Invoice").Select
    Selection.ClearContents
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "Name"
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Company Name"
    Range("C9").Select
    ActiveCell.FormulaR1C1 = "Street Address"
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "City, ST Zip Code"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = "Phone"
    Range("D10").Select
End Sub