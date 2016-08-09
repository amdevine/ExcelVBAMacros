Attribute VB_Name = "R1C1Mod"
'This macro toggles between the R1C1 reference style
'and the A1 reference style

Sub ToggleR1C1()

If Application.ReferenceStyle = xlR1C1 Then
    Application.ReferenceStyle = xlA1
    MsgBox "Reference style has been set to A1."

Else
    Application.ReferenceStyle = xlR1C1
    MsgBox "Reference style has been set to R1C1."

End If

End Sub
