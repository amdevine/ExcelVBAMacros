Attribute VB_Name = "CellFormatMod"
Option Explicit

Sub WrapText()
Attribute WrapText.VB_Description = "Keyboard shortcut to wrap text in a cell."
Attribute WrapText.VB_ProcData.VB_Invoke_Func = "W\n14"

' This macro toggles wrapping text in selected cell(s).
' Assigned keyboard shortcut Ctrl+Shift+W

    If Selection.WrapText = True Then
        Selection.WrapText = False
        MsgBox "Cells have been formatted to NOT Wrap Text."
    Else
        Selection.WrapText = True
        MsgBox "Cells have been formatted to Wrap Text."
    End If

End Sub

Sub TextFormat()
Attribute TextFormat.VB_ProcData.VB_Invoke_Func = "T\n14"

' This macro toggles selected cell(s) between Text format and General format.
' Assigned keyboard shortcut Ctrl+Shift+T
    
    If Selection.NumberFormat <> "@" Then
        Selection.NumberFormat = "@"
        MsgBox "Cells have been formatted as Text."
    
    ElseIf Selection.NumberFormat = "@" Then
        Selection.NumberFormat = "General"
        MsgBox "Cells have been formatted as General."
            
    End If

End Sub

Sub Barcode2DFormat()
Attribute Barcode2DFormat.VB_ProcData.VB_Invoke_Func = "B\n14"

' This macro toggles cell formatting to require 10 digits (to capture any leading zeroes).
' Assigned keyboard shortcut Ctrl+Shift+B
'
    If Selection.NumberFormat <> "0000000000" Then
        Selection.NumberFormat = "0000000000"
        MsgBox "Cells have been formatted to require 10 digits (Matrix 2D barcode format)."
        
    ElseIf Selection.NumberFormat = "0000000000" Then
        Selection.NumberFormat = "General"
        MsgBox "Cells have been formatted as General."
        
    End If
    
End Sub
