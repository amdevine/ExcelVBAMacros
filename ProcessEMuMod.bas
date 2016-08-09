Attribute VB_Name = "ProcessEMuMod"
Option Explicit

Sub ProcessEMuMonthly()
Attribute ProcessEMuMonthly.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ProcessEMuMonthly Macro
'

'
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "data!R1C1:R67750C23", Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:="Sheet3!R3C1", TableName:="PivotTable2", DefaultVersion _
        :=xlPivotTableVersion15
    Sheets("Sheet3").Select
    Cells(3, 1).Select
    Sheets("Sheet3").Select
    Sheets("Sheet3").name = "pivot"
    Sheets("pivot").Select
    Sheets("pivot").Move After:=Sheets(2)
    Range("A3").Select
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("CatDepartment")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("CatDivision")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields( _
        "AdmPublishWebNoPassword")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("irn"), "Count of irn", xlCount
    Range("C6").Select
    ActiveSheet.PivotTables("PivotTable2").RepeatAllLabels xlRepeatLabels
    ActiveSheet.PivotTables("PivotTable2").PivotFields("ecatalogue_key").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("irn").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("SummaryData").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields( _
        "GSFreezerProBiorepositoryID").Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("GS2DBarcode").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("NamLast").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DarCollectorNumber"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("GSFieldIDNumber"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("CatDepartment").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("CatDivision").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("AdmPublishWebNoPassword"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("GSEmbargo").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("CatObjectType").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("GSTypePrimary").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("GSTypeSecondary"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("CurCommentDate0"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DarKingdom").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DarPhylum").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DarClass").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DarOrder").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DarFamily").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DarGenus").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DarSpecies").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").RowAxisLayout xlTabularRow
End Sub
