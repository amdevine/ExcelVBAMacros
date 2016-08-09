Attribute VB_Name = "HighlightMod"
Sub YellowHighlight()
Attribute YellowHighlight.VB_Description = "Highlights selected cell(s) yellow."
Attribute YellowHighlight.VB_ProcData.VB_Invoke_Func = "Y\n14"
'
' YellowHighlight Macro
' Highlights selected cell(s) yellow.
'
' Keyboard Shortcut: Ctrl+Shift+Y
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


Sub UndoColorHighlight()
Attribute UndoColorHighlight.VB_Description = "Removes color highlighting in selected cell(s)"
Attribute UndoColorHighlight.VB_ProcData.VB_Invoke_Func = "U\n14"
'
' UndoColorHighlight Macro
' Removes color highlighting in selected cell(s)
'
' Keyboard Shortcut: Ctrl+Shift+T
'
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.ColorIndex = 1
End Sub

Sub GreyHighlight()
Attribute GreyHighlight.VB_Description = "Highlight selected cells grey and change text color to grey in order to de-emphasize selected cells."
Attribute GreyHighlight.VB_ProcData.VB_Invoke_Func = "G\n14"
'
' GreyHighlight Macro
' Changes text color to dark grey and fills cell lighter grey.
'
' Keyboard Shortcut: Ctrl+Shift+G
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ColorIndex = 15
        .PatternTintAndShade = 0
    End With
    Selection.Font.ColorIndex = 16
End Sub
