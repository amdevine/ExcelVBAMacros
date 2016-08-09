Attribute VB_Name = "CompareListsMod"
Option Explicit
Option Compare Text

Sub CompareLists()
Attribute CompareLists.VB_Description = "Compares the contents of two user-defined lists/columns/ranges and outputs the items unique to each list and in common with both lists."
Attribute CompareLists.VB_ProcData.VB_Invoke_Func = "L\n14"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This program takes two user-defined ranges as inputs and compares the items
' in those ranges, then outputs which items appear only in one range or the
' other and which items appear in both ranges.
'
' Written by Amanda Devine, 11 Aug 2015
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim rng1, rng2, listDest1, listDest2, listDest3, startList As Range
Dim list1, list2, on1not2, on2not1, on1and2, printo1n2, printo2n1, printo1a2 As Variant
Dim rowIndex1, colIndex1, rowIndex2, colIndex2 As Long
Dim r1, c1, r2, c2, dr, er, mr, d, e, m As Long
Dim stoploop As Boolean
Dim list1name, list2name As String

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get user-defined names for ranges, ranges for comparison,
' and cell where comparison lists are generated.
' Calculate what the maximum number of items are on each of these lists.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

list1name = Application.InputBox( _
            prompt:="Enter a name for your first list: ", _
            Title:="Name of First List", _
            Type:=2)

Set rng1 = Application.InputBox( _
            prompt:="Enter a range of values for " & list1name & ": ", _
            Title:="Range of " & list1name, _
            Type:=8)
            
list1 = rng1.Value
rowIndex1 = UBound(list1, 1)
colIndex1 = UBound(list1, 2)
ReDim on1not2(1 To rng1.Cells.Count + 1) ' Sets discrepancy array max size to number of items on list 1

 
list2name = Application.InputBox( _
            prompt:="Enter a name for your second list: ", _
            Title:="Name of Second List", _
            Type:=2)

Set rng2 = Application.InputBox( _
            prompt:="Enter a range of values for " & list2name & ": ", _
            Title:="Range of " & list2name, _
            Type:=8)
            
list2 = rng2.Value
rowIndex2 = UBound(list2, 1)
colIndex2 = UBound(list2, 2)
ReDim on2not1(1 To rng2.Cells.Count + 1) 'Sets discrepancy array max size to number of items on list 2


Set startList = Application.InputBox( _
            prompt:="Select where you would like your results to go: ", _
            Title:="Destination Cell", _
            Type:=8)

Set listDest1 = startList   ' Destination for first list is starting cell
Set listDest2 = Cells(startList.row, startList.Column + 1)  ' Destination for second list is 1 cell right of starting cell
Set listDest3 = Cells(startList.row, startList.Column + 2)  ' Destination for third list is 2 cells right of starting cell

ReDim on1and2(1 To Application.WorksheetFunction.Max(rng1.Cells.Count, rng2.Cells.Count) + 1) 'Sets matching list to the larger of list 1 or 2


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Loop through list 1, comparing each item to every item in list 2
' until it finds or does not find a match. If a match is found, adds
' that item to the on1and2 list. If no match is found, adds that item
' to the on1not2 list.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

dr = 1  ' Row counter for on1not2 list
mr = 1  ' Row counter for on1and2 list

on1not2(dr) = "Items in " & list1name & " that are not in " & list2name & ":"
on1and2(mr) = "Items that appear in both lists:"

For r1 = 1 To rowIndex1         'Increments list 1 row
    For c1 = 1 To colIndex1     'Increments list 1 column
                   
            For r2 = 1 To rowIndex2     'Increments list 2 row
                For c2 = 1 To colIndex2 'Increments list 2 column
                    
                    
                        If list1(r1, c1) = list2(r2, c2) Then   'If items match
                            mr = mr + 1
                            on1and2(mr) = list1(r1, c1)
                            stoploop = True
                            Exit For    'Stops iterating through list 2 items in that row
                        ElseIf list1(r1, c1) <> list2(r2, c2) And r2 = rowIndex2 And c2 = colIndex2 Then 'If items do not match, add to discrepancy list.
                            dr = dr + 1
                            on1not2(dr) = list1(r1, c1)
                        End If
    
                Next c2 'Progresses to next item in list 2 row if no match
                
                If stoploop = True Then
                    Exit For    'Stops iterating through list 2 rows i.e. stops iterating through list 2
                End If
                                
            Next r2     'Progresses to next list 2 row if no match
            
            stoploop = False
            
    Next c1     'Progression to next item in that list 1 row
Next r1         'Progression to next row in list 1

ReDim Preserve on1not2(1 To dr)     'Redimensions the array based on the number of items in it.
ReDim Preserve on1and2(1 To mr)     'Redimensions the array based on the number of items in it.


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Loop through list 2, comparing each item to every item in list 1
' until it finds or does not find a match. If a match is found, nothing happens
' (item already added from list 1 looping). If no match is found, adds that item
' to the on1not2 list.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

er = 1  ' on2not1 row counter
on2not1(er) = "Items in " & list2name & " that are not in " & list1name & ":"

For r2 = 1 To rowIndex2         'Increments list 2 row
    For c2 = 1 To colIndex2     'Increments list 2 column
                   
            For r1 = 1 To rowIndex1     'Increments list 1 row
                For c1 = 1 To colIndex1 'Increments list 1 column
                    
                    
                        If list2(r2, c2) = list1(r1, c1) Then   'If items match
                            stoploop = True
                            Exit For    'Stops iterating through list 2 items in that row
                        ElseIf list2(r2, c2) <> list1(r1, c1) And r1 = rowIndex1 And c1 = colIndex1 Then 'If items do not match, add to discrepancy list.
                            er = er + 1
                            on2not1(er) = list2(r2, c2)
                        End If
    
                Next c1 'Progresses to next item in list 1 row if no match
                
                If stoploop = True Then
                    Exit For    'Stops iterating through list 1 rows i.e. stops iterating through list 1
                End If
                                
            Next r1     'Progresses to next list 1 row if no match
            
            stoploop = False
            
    Next c2     'Progression to next item in that list 2 row
Next r2         'Progression to next row in list 2

ReDim Preserve on2not1(1 To er)     'Resize array to only be the number of items


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Print the discrepancy lists and the match list at the location
' specified by the user. Values print in three columns:
' Column 1 = Values unique to the first list
' Column 2 = Values unique to the second list
' Column 3 = Values matching on both lists
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ReDim printo1n2(1 To dr, 1 To 1) ' Converting on1not2 to column
For d = 1 To dr
    printo1n2(d, 1) = on1not2(d)
Next d

ReDim printo2n1(1 To er, 1 To 1) ' Converting on2not1 to column
For e = 1 To er
    printo2n1(e, 1) = on2not1(e)
Next e

ReDim printo1a2(1 To mr, 1 To 1) ' Converting on1and2 to column
For m = 1 To mr
    printo1a2(m, 1) = on1and2(m)
Next m

listDest1.Resize(UBound(on1not2), 1).Value = printo1n2

listDest2.Resize(UBound(on2not1), 1).Value = printo2n1

listDest3.Resize(UBound(on1and2), 1).Value = printo1a2


End Sub
