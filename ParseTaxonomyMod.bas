Attribute VB_Name = "ParseTaxonomyMod"
Option Explicit

Sub ParseTax()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This program was written to format taxonomic output from the Open Tree
' of Life. Raw data were exported from Access - each row is a unique genus,
' with all of the names and ranks of parent taxonomic groups listed to the
' left of the genus. This program takes those varied parents and aligns them
' by taxonomic level.
'
' While this program was written specifically for this one output file, if
' some adaptations are made, it could be a general program for formatting all
' taxonomic data that are listed in this format.
'
'
' Written by Amanda Devine, 18 Nov 2015
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' As written, both raw data and output hierarchy must be in table form,
' and the output hierarchy will be sorted in the same order as the raw data.
' Required headers for the output hierarchy table:
'        domain
'        kingdom
'        subkingdom
'        infrakingdom
'        superphylum
'        phylum
'        subphylum
'        infraphylum
'        parvphylum
'        superclass
'        class
'        subclass
'        infraclass
'        superorder
'        order
'        suborder
'        superfamily
'        family
'        subfamily
'        genus
' I also added a "genus_id" column at the end and copy and pasted the values
' from Raw to Hierarchy.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim row, rcol, numRows, numCols As Integer
Dim htab, rtab As ListObject
Dim name As String

' htab = final hierarchy table, rtab = raw data table
' CHANGE THESE depending on sheets/tables you're working with
Set htab = Worksheets("Hierarchy").Range("Hierarchy").ListObject
Set rtab = Worksheets("RawSorted").Range("Raw").ListObject

' The number of rows in the raw data table (headers not included)
' and the number of columns (IDs included)
numRows = rtab.DataBodyRange.Rows.Count
numCols = rtab.DataBodyRange.Columns.Count

' Iterates through every column in a row, and all rows in the raw data table
For row = 1 To numRows
    For rcol = 2 To numCols

        ' The taxonomic name is equal to the column to the left of the taxonomic rank
        name = rtab.ListColumns(rcol - 1).DataBodyRange(row)
        
        ' Switch statement evaluating rank of parent and assigning name of parent
        ' to appropriate hierarchy location
        Select Case rtab.ListColumns(rcol).DataBodyRange(row)
        
            Case "domain"
                htab.ListColumns("domain").DataBodyRange(row) = name
        
            Case "kingdom"
                htab.ListColumns("kingdom").DataBodyRange(row) = name
        
            Case "subkingdom"
                htab.ListColumns("subkingdom").DataBodyRange(row) = name
        
            Case "infrakingdom"
                htab.ListColumns("infrakingdom").DataBodyRange(row) = name
        
            Case "superphylum"
                htab.ListColumns("superphylum").DataBodyRange(row) = name
        
            Case "phylum"
                htab.ListColumns("phylum").DataBodyRange(row) = name
        
            Case "subphylum"
                htab.ListColumns("subphylum").DataBodyRange(row) = name
        
            Case "infraphylum"
                htab.ListColumns("infraphylum").DataBodyRange(row) = name
        
            Case "parvphylum"
                htab.ListColumns("parvphylum").DataBodyRange(row) = name
        
            Case "superclass"
                htab.ListColumns("superclass").DataBodyRange(row) = name
        
            Case "class"
                htab.ListColumns("class").DataBodyRange(row) = name
        
            Case "subclass"
                htab.ListColumns("subclass").DataBodyRange(row) = name
        
            Case "infraclass"
                htab.ListColumns("infraclass").DataBodyRange(row) = name
        
            Case "superorder"
                htab.ListColumns("superorder").DataBodyRange(row) = name
            
            Case "order"
                htab.ListColumns("order").DataBodyRange(row) = name
                
            Case "suborder"
                htab.ListColumns("suborder").DataBodyRange(row) = name
                
            Case "superfamily"
                htab.ListColumns("superfamily").DataBodyRange(row) = name
        
            Case "family"
                htab.ListColumns("family").DataBodyRange(row) = name
            
            Case "subfamily"
                htab.ListColumns("subfamily").DataBodyRange(row) = name
                
            Case "genus"
                htab.ListColumns("genus").DataBodyRange(row) = name
        
        End Select
        
            
    Next rcol
Next row


End Sub


