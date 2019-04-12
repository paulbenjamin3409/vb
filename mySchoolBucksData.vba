' takes pasted data from the transactions page and transforms it for table use
' Use:
' 1. Copy transactions page table (header and values) and paste into A1 of excel spreadsheet
' 2. Copy vba into Marco/Button
' 3. Run.
Sub FromColumnToRowData()
    Dim i As Integer
    Dim o As Integer
    Dim lastrow As Long
    
    o = 1
    i = 1
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    Do While i < lastrow
        Dim r1 As String
        r1 = "A" & i
        Range(r1).Select
        Range("A" & i & ":A" & i + 6).Select
        Selection.Copy
        Range("B" & o).Select
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
        i = i + 7
        o = o + 1
    Loop

End Sub
