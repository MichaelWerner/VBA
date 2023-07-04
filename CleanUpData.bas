Sub CleanUpData()
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim rngUsed As Range, Cell As Range
    Dim arrUnwantedASCIICodes
    Dim i As Integer, j As Integer
    
    arrUnwantedASCIICodes = Array(127, 129, 141, 143, 144, 157, 160)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set WB = ActiveWorkbook
    
    For j = 1 To WB.Sheets.Count
        
        Set WS = WB.Sheets(j)
        On Error Resume Next
        'xlConstants: Do not look for cells having a formula
        'xlNumbers + xlTextValues: in xlConstants look for numbers and text only (should be all)
        Set rngUsed = WS.UsedRange.SpecialCells(xlConstants, xlNumbers + xlTextValues)
        If rngUsed Is Nothing Then
            Exit Sub
        End If
        On Error GoTo 0
        
        
        For Each Cell In rngUsed
            'Clean: removes ASCII Code 0 - 31 from the cell
            'Trim:  removes leading and trailing blanks (ASCII 32)
            Cell = Trim(WorksheetFunction.Clean(Cell))
            
            'Now remove all unwanted ASCII Codes
            For i = LBound(arrUnwantedASCIICodes) To UBound(arrUnwantedASCIICodes)
                Cell = Replace(Cell, Chr(arrUnwantedASCIICodes(i)), "")
            Next i
        Next Cell
    Next j
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
