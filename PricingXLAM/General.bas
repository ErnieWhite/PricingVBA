Attribute VB_Name = "General"
Option Explicit
' General functions handy for pricing

Sub AddNetColumn()
    ' Add a NET column
    '
    ' Inserts a column at the current cell location
    ' If the column can not be inserted, displays an error message and exits
    ' The first row is treated as a header row.
    ' The header is created by taking the value from the column to the right and adding the suffix BASIS to it. Separated by a space
    '
    ' Keyboard Shortcut: Ctrl+Shift+n
    '
    
    ' get a handle to the current worksheet.
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Get the current column number
    Dim insertAtColumnNum As Long
    insertAtColumnNum = Application.ActiveCell.Column
    
    On Error GoTo ErrorHandler ' Redirects to the error handler if an error occurs
    
    ws.Range(ws.Cells(1, insertAtColumnNum), ws.Cells(1, insertAtColumnNum)).EntireColumn.Insert shift:=xlToRight ' insert a column at the active cell
    ws.Range(ws.Cells(1, insertAtColumnNum), ws.Cells(1, insertAtColumnNum)) = ws.Range(ws.Cells(1, insertAtColumnNum + 1), ws.Cells(1, insertAtColumnNum + 1)).Value2 & " BASIS" ' set the column header
    
    Dim lastRow As Long
    lastRow = ws.UsedRange.Rows.Count ' find the last used row on the worksheet
    
    Dim data As Variant
    data = ws.Range(ws.Cells(2, insertAtColumnNum + 1), ws.Cells(lastRow, insertAtColumnNum + 1)).Value2 ' get a copy of the data one column to the right
    
    Dim d As Variant
    Dim i As Long
    For i = LBound(data) To UBound(data)
        If data(i, 1) <> "" Then
            data(i, 1) = "NET"
        End If
    Next i
    
    ws.Range(ws.Cells(2, insertAtColumnNum), ws.Cells(lastRow, insertAtColumnNum)) = data
    
    On Error GoTo 0 ' turn error handling back off, if an untrapped error happens somewhere else we don't want to end up back here
    
    Exit Sub ' Ensures that the error handler runs only when an error occurs
    
ErrorHandler:
    MsgBox "An error occured: " & Err.Description, vbCritical, "Error #" & Err.Number
    Resume Next ' continues with the next line
    Err.Clear ' clears the error
    
End Sub

Sub SetTheDollarFormat()
    ' Formats the active cells with "$###0.000"
    '
    ' Keyboard Shortcut: Ctrl+Shift+f
    '
    Selection.NumberFormat = "$###0.000"
    
End Sub

Sub MoveToEnd()
    '
    ' Moves the column containing the active cell to the right side of the data based on the 1st row
    '
    ' Keyboard Shortcut: Ctrl+Shift+e
    '
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim moveFrom As Long
    moveFrom = ActiveCell.Column
    
    Dim moveTo As Long
    moveTo = ws.Cells(1, ws.Cells.Columns.Count).End(xlToLeft).Column + 1
    
    ws.Range(ws.Cells(1, moveFrom), ws.Cells(1, moveFrom)).EntireColumn.Select
    Selection.Cut
    ws.Range(ws.Cells(1, moveTo), ws.Cells(1, moveTo)).Insert shift:=xlToRight
    
    On Error GoTo 0
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occured: " & Err.Description, vbCritical, "Error #" & Err.Number
    Err.Clear ' clears the error
    Resume Next ' continues with next line
    
End Sub
