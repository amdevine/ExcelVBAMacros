Attribute VB_Name = "AutofillMod"
Option Explicit

'This macro toggles autofilling table columns with formulas on and off.

Sub ToggleAutofillTable()

If Application.AutoCorrect.AutoFillFormulasInLists = True Then
    Application.AutoCorrect.AutoFillFormulasInLists = False
    MsgBox "Table columns have been set to NOT autofill."

Else
    Application.AutoCorrect.AutoFillFormulasInLists = True
    MsgBox "Table columsn have been set TO autofill."

End If

End Sub
