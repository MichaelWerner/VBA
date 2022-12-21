Attribute VB_Name = "Module1"
Option Explicit
Sub changeColorForPartOfCell()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sSearch As String, sLongSearch As String
    
    Dim iFirstRow As Integer, iLastRow As Integer
    Dim i As Integer
    
    Set wb = Application.ActiveWorkbook
    Set ws = wb.ActiveSheet
    
    iFirstRow = 2
    iLastRow = ws.UsedRange.Rows.Count
    
    sSearch = "branch_office_managers"
    sLongSearch = sSearch & "@zwickerrpc.com"
    
    For i = iFirstRow To iLastRow
        'MsgBox InStr(1, ws.Cells(i, 3).Value, sSearch)
        If InStr(1, ws.Cells(i, 3).Value, sSearch) > 0 Then
        'green:RGB(0, 150, 25)
        'red:RGB(255, 0, 0)
        'blue:RGB(0, 0, 255)
            ws.Cells(i, 3).Characters(InStr(1, ws.Cells(i, 3).Value, sSearch), Len(sLongSearch)).Font.Color = RGB(0, 0, 255)
        End If
    Next i
End Sub
