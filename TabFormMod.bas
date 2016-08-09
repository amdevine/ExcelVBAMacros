Attribute VB_Name = "TabFormMod"
Option Explicit

Sub TableFormatting()
Attribute TableFormatting.VB_Description = "Formats selected table by removing color and striping, bolding column headers, and wrapping column headers."
Attribute TableFormatting.VB_ProcData.VB_Invoke_Func = "R\n14"

    Dim SelectedCell As Range
    Dim TableName As String
    Dim ActiveTable As ListObject

    Set SelectedCell = ActiveCell

    'Determine if ActiveCell is inside a Table
    On Error GoTo NoTableSelected
    
    TableName = SelectedCell.ListObject.name
    Set ActiveTable = ActiveSheet.ListObjects(TableName)
    
    On Error GoTo 0

    'Do something with your table variable (ie Add a row to the bottom of the ActiveTable)
    ActiveTable.TableStyle = ""
    ActiveTable.HeaderRowRange.Font.Bold = True
    ActiveTable.HeaderRowRange.WrapText = True
   
  
    Exit Sub

'Error Handling
    
NoTableSelected:
    
    MsgBox "There is no Table currently selected!", vbCritical
        
End Sub


Sub PivotFormatting()

    Dim SelectedCell As Range
    Dim PivotName As String
    Dim ActivePivot As PivotTable
    Dim pf As PivotField
    Dim pf2 As PivotField

    Set SelectedCell = ActiveCell

    'Determine if ActiveCell is inside a Table
    On Error GoTo NoPivotSelected
    
    PivotName = ActiveCell.PivotTable.name
    Set ActivePivot = ActiveSheet.PivotTables(PivotName)
    
    On Error GoTo 0

    'Format Pivot Table to be tabular, to repeat labels, to not show +/- button, to have no color formatting
    With ActivePivot
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .ShowDrillIndicators = False
        .TableStyle2 = ""
        .ColumnGrand = False
        .RowGrand = False
        .DataLabelRange.Font.Bold = True
    End With
    
    'Format each visible Pivot Field to have a bold header and not show a subtotal
    For Each pf In ActivePivot.PivotFields
        If pf.Orientation = 1 Then
            pf.LabelRange.Select
            Selection.Font.Bold = True
            pf.Subtotals(1) = False
        End If
        Next pf
        
    
    Exit Sub
    

'Error Handling
    
NoPivotSelected:
        
    MsgBox "There is no Pivot Table currently selected!", vbCritical

End Sub
