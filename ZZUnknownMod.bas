Attribute VB_Name = "ZZUnknownMod"
Option Explicit

Sub ZZUnknown()
Attribute ZZUnknown.VB_Description = "Fills in selected cell(s) with ""ZZ_unknown""."
Attribute ZZUnknown.VB_ProcData.VB_Invoke_Func = "Z\n14"
'
' ZZUnknown Macro
' Fills "ZZ_unknown" into selected cell(s).
'
' Keyboard Shortcut: Ctrl+Shift+Z
'
    Selection.FormulaR1C1 = "ZZ_unknown"
    
End Sub
