Attribute VB_Name = "GlblSetup"
Option Explicit

Sub GlblAddColumns(ws As Worksheet)
    '
    ' Adds all the columns needed for the PRC_UPDATE_GLBL import map
    '
    
    ' add the columns that are actually used for the import file (blue)
    ws.Range("AF1:AN1").EntireColumn.Insert Shift:=xlToRight
    
    ' add the informational(yellow) colums
    ws.Range("AO1:AR1").EntireColumn.Insert Shift:=xlToRight
    
    ' add the cleaned catalog number column (yellow)
    ws.Range("e1").EntireColumn.Insert Shift:=xlToRight
    
End Sub

Sub GlblSetColumnHeaders(ws As Worksheet)
    '
    ' sets the column headers for the columns we added
    '
    
    ' Import data headers
    ws.Range("AG1:AO1") = Array("MSC UNIQUE", "LIST PRICE", "MULTIPLIER", "REP COST", "EFF DATE", "UMRP", "STANDARD COST", "DC COST", "CMP")
    
    ' Variance column headers
    ws.Range("AP1:AS1") = Array("CMP Margin", "LIST Var", "REP Var", "CMP Var")
    
    ' Clean catalog header
    ws.Range("E1") = "Cleaned Catalog"
    
End Sub

Sub GlblSetHeaderColorCodes(ws As Worksheet)
    '
    ' Sets the backround color coding for the header row
    '
    
    ' the import import fields are blue
    ws.Range("A1:AF1", "AS1:CO1").Interior.Color = RGB(247, 150, 70) ' Orange
    ws.Range("AG1:AO1").Interior.Color = RGB(0, 0, 255) ' blue
    ws.Range("AG1:AO1").Font.Color = RGB(255, 255, 255) ' White
    ws.Range("AP1:AS1").Interior.Color = RGB(255, 255, 0) ' yellow
    ws.Range("E1:E1").Interior.Color = RGB(255, 255, 0) ' Yellow
    
End Sub

Sub GlblCleanedFormula(ws As Worksheet)
    '
    ' enters the catalog
    '
    ' TODO: change this to a static value
    ws.Range("E2:E" & ws.UsedRange.Rows.Count).Formula = "=LET(MODEL,INDIRECT(""rc[-1]"",FALSE),SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(CLEAN(MODEL),""-"",""""),""."",""""),""/"",""""),""\"",""""),""_"",""""),"","",""""), "" "",""""))"

End Sub

Sub GlblCopyMscUnique(ws As Worksheet)
    '
    ' copies the values from the msc unique in column A to column AG
    '
    
    ws.Range("AG2:AG" & ws.UsedRange.Rows.Count) = ws.Range("A2:A" & ws.UsedRange.Rows.Count).Value2
End Sub

Sub GlblVarianceFormulas(ws As Worksheet)
    '
    ' sets up the 3 variance formulas, LIST Var, REP Var, CMP Var
    '
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' subtract 1 since this gives us one row past the last last row with data
    
    Dim data As Variant
    
    Dim i As Long
    
    ' LIST Variance formulas
    ' ws.Range("AP2:AP" & lastRow).FormulaR1C1 = "=1-INDIRECT(""RC[-13]"", FALSE)/INDIRECT(""RC[-8]"", FALSE)"
    data = ws.Range("AQ2:AQ" & lastRow).Value2
    For i = 1 To lastRow - 1 ' the data file is from 2 to lastrow. we are starting at 1 so take one from lastrow
        data(i, 1) = "=1-AC" & i + 1 & "/AH" & i + 1
     Next i
    ws.Range("AQ2:AQ" & lastRow).Formula2 = data
    
    ' REP Variance
    'ws.Range("AQ2:AQ" & lastRow).FormulaR1C1 = "=1-INDIRECT(""RC[-12]"", FALSE)/INDIRECT(""RC[-7]"", FALSE)"
    data = ws.Range("AR2:AR" & lastRow).Value2
    For i = 1 To lastRow - 1
        data(i, 1) = "=1-AE" & i + 1 & "/AJ" & i + 1
    Next i
    ws.Range("AR2:AR" & lastRow).Formula2 = data
    
    ' CMP Variance
    'ws.Range("AR2:AR" & lastRow).FormulaR1C1 = "=1-INDIRECT(""RC[-12]"", FALSE)/INDIRECT(""RC[-3]"", FALSE)"
    data = ws.Range("AS2:AS" & lastRow).Value2
    For i = 1 To lastRow - 1
        data(i, 1) = "=1-AF" & i + 1 & "/AO" & i + 1
    Next i
    ws.Range("AS2:AS" & lastRow).Formula2 = data
    
End Sub

Sub GlblCMPFormula(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim data As Variant
    data = Range("AO2:AO" & lastRow).Value2
    Dim i As Long
    For i = 2 To lastRow
        data(i - 1, 1) = "=LET(CMP,AF" & i & " / (1 - AQ" & i & "), IF(CMP < AH" & i & ", CMP, AH" & i & "))"
    Next i
    ws.Range("AO2:AO" & lastRow).Formula2 = data
End Sub

Sub GlblCMPMarginFormula(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim data As Variant
    data = Range("AP2:AP" & lastRow).Value2
    Dim i As Long
    For i = 2 To lastRow
        data(i - 1, 1) = "=1 - AJ" & i & " / AO" & i
    Next i
    ws.Range("AP2:AP" & lastRow).Formula2 = data
End Sub

Sub GlblFormatColumns(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim rng As Range
    
    ' format the list price column
    Range("AH2:AH" & lastRow).NumberFormat = "0.000"
    Range("AI2:AI" & lastRow).NumberFormat = "0.0000"
    Range("AJ2:AJ" & lastRow).NumberFormat = "0.000"
    Range("AL2:AO" & lastRow).NumberFormat = "0.000"
    Range("AP2:AS" & lastRow).NumberFormat = "0%;[Red]-0%"
    
    Selection.NumberFormat = "0.000"
End Sub

Sub ConvertPriceSheetDataToNumbers(ws As Worksheet)
'
' Macro2 Macro
'

'
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim data As Variant
    data = Range("AC2:AF" & lastRow).Value2
        
    Dim i As Long
    Dim j As Long
    For i = LBound(data, 1) To UBound(data, 1)
        For j = LBound(data, 2) To UBound(data, 2)
            data(i, j) = Val(data(i, j))
        Next j
    Next i
    
    Range("AC2:AF" & lastRow) = data
    
End Sub


Sub GlblAutoFilter(ws As Worksheet)
'
' Macro3 Macro
'

'
    If Not ws.AutoFilterMode Then
        ws.Range("A1").AutoFilter
    End If
    
End Sub

Sub GlblFreezeTopRow(aw As Window)
'
' GlblFreezeTopRow Macro
'

'
    
    With aw
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    
End Sub

Sub GlblTurnOffWordWrap(ws As Worksheet)
'
' GlblTurnOfWordWrap Macro
'

'
    ws.Cells.WrapText = False

End Sub

Sub GlblAutoSize()
'
' GlblAutoSize Macro
'

'
    Cells.Select.EntireColumn.AutoFit
    
End Sub

Sub GlblVarianceReports(ws As Worksheet)

    ws.Range("CQ1").Formula2 = "= ""LIST Average Var: "" & text(subtotal(101,AQ:AQ), ""0.00%"")"
    ws.Range("CR1").Formula2 = "= ""REP Average Var: "" & text(subtotal(101, AR:AR), ""0.00%"")"
    ws.Range("CS1").Formula2 = "= ""CMP Average Var: "" & text(subtotal(101, AS:AS), ""0.00%"")"
    
End Sub

Sub GlblGroupColumns(ws As Worksheet)
    ws.Columns("F:AB").Group
    ws.Columns("AT:CM").Group
    ws.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
End Sub

Sub turn_off()
'
' turn_off Macro
'

' settings for speeding up the sheet setup
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .StatusBar = "Setting up worksheet..."
    End With
End Sub

Sub turn_on()
'
' turn_on Macro
'

' undo setting for speeding up the sheet setup

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .StatusBar = ""
    End With
End Sub

Sub Setup_PDW2Extract_For_GLBL()
Attribute Setup_PDW2Extract_For_GLBL.VB_Description = "Sets up the PDWv2Extract for doing a PRC_UPDATE_GLBL import"
Attribute Setup_PDW2Extract_For_GLBL.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Setup_PDW2Extract_For_GLBL Macro
' Sets up the PDWv2Extract for doing a PRC_UPDATE_GLBL import
'

'
    Dim aw As Window
    Set aw = ActiveWindow
    
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    On Error GoTo CLEANEXIT
    Call turn_off
    
    GlblAddColumns ws
    GlblSetColumnHeaders ws
    GlblSetHeaderColorCodes ws
    GlblCleanedFormula ws
    GlblCopyMscUnique ws
    GlblVarianceFormulas ws
    GlblFreezeTopRow aw
    GlblAutoFilter ws
    GlblTurnOffWordWrap ws
    GlblVarianceReports ws
    GlblGroupColumns ws
    GlblCMPFormula ws
    GlblFormatColumns ws
    GlblCMPMarginFormula ws
    ConvertPriceSheetDataToNumbers ws
    
    ws.Cells.EntireColumn.AutoFit
    
    ws.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    
    ws.Range("AG2").Select
    
CLEANEXIT:
    Call turn_on
    On Error GoTo 0
End Sub


Sub testGlblAddColumns()
    
    GlblAddColumns ActiveSheet
    
End Sub

Sub testGlblSetColumnHeaders()
    
    GlblSetColumnHeaders ActiveSheet
    
End Sub

Sub testGlblSetHeaderColorCodes()
    
    GlblSetHeaderColorCodes ActiveSheet
    
End Sub

Sub testGlblCleanedFormula()
    
    GlblCleanedFormula ActiveSheet
    
End Sub

Sub testGlblCopyMscUnique()
    
     GlblCopyMscUnique ActiveSheet
    
End Sub

Sub testGlblVarianceFormulas()

    GlblVarianceFormulas ActiveSheet
    
End Sub

Sub testGlblVarianceReports()

    
    GlblVarianceReports ActiveSheet
    
End Sub

Sub testGlblGroupColumns()

    GlblGroupColumns ActiveSheet
    
End Sub

