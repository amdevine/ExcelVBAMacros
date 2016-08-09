Attribute VB_Name = "FillWellsMod"
'This module will place a list of well locations (e.g. C07)
'in a column in a spreadsheet.


Sub PlateWellLocations()

'Capturing the coordinates of the selected cell(s)
actRow = ActiveCell.row
actCol = ActiveCell.Column

'Defining the position of the cell and the column/row names
well = 1
numString = ""
letString = ""

'Inserts the word "Well" at the top of the list
Cells(actRow, actCol) = "Well"

'The column location
For numLoc = 1 To 12

    'The row location
    For letLoc = 65 To 72
    
        'Converts column location to string and row location to alphanumeric character
        numString = CStr(numLoc)
        letString = Chr(letLoc)
    
        'If column is less than 10, adds a "0" before the column
        If numLoc < 10 Then
            Cells(actRow + well, actCol) = letString + CStr(0) + numString
            
        'If column is 10 or more, does not add a "0" before the column
        ElseIf numLoc >= 10 Then
            Cells(actRow + well, actCol) = letString + numString
            
        End If
                
        'Increments the cell down 1
        well = well + 1
        
    Next letLoc
    
Next numLoc
    
    
    
End Sub


Sub FreezerProWellLocations()
Attribute FreezerProWellLocations.VB_Description = "Inserts a column of 96 well locations formatted for FreezerPro uploads (A/1, B/1, C/1, etc.)"
Attribute FreezerProWellLocations.VB_ProcData.VB_Invoke_Func = "F\n14"

'Capturing the coordinates of the selected cell(s)
actRow = ActiveCell.row
actCol = ActiveCell.Column

'Defining the position of the cell and the column/row names
well = 0
numString = ""
letString = ""

'The column location
For numLoc = 1 To 12

    'The row location
    For letLoc = 65 To 72
    
        'Converts column location to string and row location to alphanumeric character
        numString = CStr(numLoc)
        letString = Chr(letLoc)
    
        Cells(actRow + well, actCol) = letString + "/" + numString
                            
        'Increments the cell down 1
        well = well + 1
        
    Next letLoc
    
Next numLoc
    
    
    
End Sub

