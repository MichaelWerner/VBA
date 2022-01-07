Attribute VB_Name = "Module1"
Option Explicit
Sub test()

    ActiveWorkbook.Queries.Add Name:="getData", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Csv.Document(File.Contents(""C:\Users\micha\Documents\test\test.csv""),[Delimiter="","", Columns=16, Encoding=1252, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}, {""Column5"", type text}, {""Column6"", " & _
        "type text}, {""Column7"", type text}, {""Column8"", type text}, {""Column9"", type text}, {""Column10"", type text}, {""Column11"", type text}, {""Column12"", type text}, {""Column13"", type text}, {""Column14"", type text}, {""Column15"", type text}, {""Column16"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""getData"";Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM getData")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "getData"
        .Refresh BackgroundQuery:=False
    End With
    ActiveSheet.ListObjects("getData").ShowTableStyleRowStripes = False
    ActiveSheet.ListObjects("getData").ShowHeaders = False
    ActiveSheet.Name = "getData"
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    
    Application.Sheets("Data").Activate
    ActiveSheet.Range("A1").Select
    ActiveCell.FormulaR1C1 = "=getData[@Column3]"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=getData[@Column11]"
    Range("A1:B1").Select
    Selection.AutoFill Destination:=Range("A1:B2"), Type:=xlFillDefault
    Range("A1:B2").Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    
    Application.DisplayAlerts = False
    ActiveWorkbook.Queries("getData").Delete
    Application.Sheets("getData").Delete
    ActiveSheet.Shapes.Range(Array("Button 4")).Delete
    Application.ActiveWorkbook.SaveAs "c:\users\micha\documents\test\t1.xlsx", xlWorkbookDefault
    Application.DisplayAlerts = True
End Sub

