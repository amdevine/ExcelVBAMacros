Attribute VB_Name = "PrintBRLabelsMod"
Sub PrintBRLabels()
Attribute PrintBRLabels.VB_Description = "A1: ""Biorepository Number"". A2-...: Biorepository Numbers. Formats a column of BR Numbers into a format for printing labels."
Attribute PrintBRLabels.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PrintBRLabels Macro
' Takes a "Field_sheet_for_collector" report from FreezerPro and formats it for printing Biorepository Labels using LabelMark 5 software.
'


' Declaring variable to determine the row of the last Biorepository Number
    Dim lastRow As String

' Delete Barcode number column from sheet
    Columns(1).EntireColumn.Delete

' Clears extraneous headers from Field Sheet
    Range("C1:H1").Clear
      
' Changes "Biorepository Number" to "BR NUMBER"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "BR NUMBER"
    
' Splits BR NUMBER values into two columns for CAP 1 and CAP 2
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    lastRow = CStr(Selection.Rows.Count + 1)
    Selection.TextToColumns Destination:=Range("B2"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 2), Array(4, 2)), TrailingMinusNumbers:=True
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "CAP 1"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "CAP 2"
    
' Creates column for LINE 2
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "LINE 2"
    
' Fills NMNH BIOREPOSITORY for each sample
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "NMNH BIOREPOSITORY"
    Range("D2:D" & lastRow).Select
    Selection.FillDown
    
' AutoFits column width for all cells
    Cells.Select
    Selection.Columns.AutoFit


' Macro-generated code to sort list in descending BR number. Needs to be figured out.
'Sub SortingMacrooo()
''
'' SortingMacrooo Macro
''
'
''
'    ActiveWorkbook.Worksheets("vials_report_1441904575_621").Sort.SortFields.Clear
'    ActiveWorkbook.Worksheets("vials_report_1441904575_621").Sort.SortFields.Add _
'        Key:=Range("A2:A101"), SortOn:=xlSortOnValues, Order:=xlDescending, _
'        DataOption:=xlSortNormal
'    With ActiveWorkbook.Worksheets("vials_report_1441904575_621").Sort
'        .SetRange Range("A1:D101")
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'End Sub


End Sub
