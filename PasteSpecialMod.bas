Attribute VB_Name = "PasteSpecialMod"
Option Explicit

Sub PasteSpecial()
Attribute PasteSpecial.VB_Description = "Paste Special -> Just Values"
Attribute PasteSpecial.VB_ProcData.VB_Invoke_Func = "V\n14"
'
' Keyboard shortcut to Paste Special -> Values
' Ctrl+Shift+V
'
'
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub
