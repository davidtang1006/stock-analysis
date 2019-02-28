Attribute VB_Name = "Module2"
Sub data_enquiry_02()
'
' data_enquiry_02 巨集
'

' Variables
Dim code, tname_01, tname_02 As String
Dim svalue, evalue, ccount_01, ccount_02, rcount_01, rcount_02 As Long
Dim i, j, k As Integer

' Count the number of stocks
rcount_01 = Cells(Rows.Count, "A").End(xlUp).Row - 1

' Set the value of ccount_01
ccount_01 = 8

' Delete the original tables
' ccount_02 = Cells(2, Columns.Count).End(xlToLeft).Column - 10 + 1
ccount_02 = 115
If ccount_02 >= 1 Then
    ' Columns("J:P").ClearContents
    ' Columns("J:P").ClearFormats
    Columns(10).Resize(, ccount_02).ClearContents
    Columns(10).Resize(, ccount_02).ClearFormats
    Columns(10).Resize(, ccount_02).ColumnWidth = 8.44
End If

' Delete the original charts
If ActiveSheet.ChartObjects.Count > 0 Then
    ActiveSheet.ChartObjects.Delete
End If

' Start the For loop
For i = 1 To rcount_01

' Calculate the value of svalue and evalue, etc.
code = Cells(i + 1, 1)
tname_01 = "Table" & " " & i
tname_02 = "Table_" & i
svalue = 1451577600 + DateDiff("d", "1/1/2016", Date) * 86400 - 364 * 86400
evalue = 1451577600 + DateDiff("d", "1/1/2016", Date) * 86400

' Clear the original queries
For Each qr In ThisWorkbook.Queries
    If qr = tname_01 Then
       qr.Delete
    End If
Next qr

' Clear the existing connections
For Each cn In ThisWorkbook.Connections
    cn.Delete
Next cn

    ' Import the data
    ActiveWorkbook.Queries.Add Name:=tname_01, Formula:= _
        "let" & Chr(13) & "" & Chr(10) & _
        "來源 = Web.Page(Web.Contents(""https://hk.finance.yahoo.com/quote/" & code & ".HK/history?period1=" & svalue & "&period2=" & evalue & "&interval=1d&filter=history&frequency=1d""))," & _
        Chr(13) & "" & Chr(10) & "Data2 = 來源{2}[Data]," & Chr(13) & "" & Chr(10) & _
        "已變更類型 = Table.TransformColumnTypes(Data2,{{""日期"", type date}, {""開市"", type number}, {""最高"", type number}, {""最低"", type number}, {""收市*"", type number}, {""經調整收市價" & _
        "**"", type number}, {""成交量"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "已變更類型" & ""
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""" & tname_01 & """", _
        Destination:=Cells(2, 10 + (i - 1) * (ccount_01 + 1))).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & tname_01 & "]")
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
        .ListObject.DisplayName = tname_02
        .Refresh BackgroundQuery:=False
    End With
    
    ' Tweaking
    Range(tname_02 & "[[#Headers],[收市*]]").Select
    ActiveCell.FormulaR1C1 = "收市"
    Range(tname_02 & "[[#Headers],[經調整收市價**]]").Select
    ActiveCell.FormulaR1C1 = "經調整收市價"
    Range(tname_02).Select
    ActiveSheet.ListObjects(tname_02).ShowAutoFilterDropDown = True
    For j = 10 + (i - 1) * (ccount_01 + 1) To 10 + (i - 1) * (ccount_01 + 1) + 6
        Columns(j).Select
        Selection.ColumnWidth = Selection.ColumnWidth * 1.5
    Next j
    
    ' Sorting
    Range(tname_02 & "[[#Headers],[日期]]").Select
    ActiveWorkbook.Worksheets("主頁面").ListObjects(tname_02).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("主頁面").ListObjects(tname_02).Sort.SortFields.Add _
        Key:=Range(tname_02 & "[[#Headers],[日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("主頁面").ListObjects(tname_02).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Add the headers
    Range(Cells(1, 10 + (i - 1) * (ccount_01 + 1)), Cells(1, 10 + (i - 1) * (ccount_01 + 1) + 6)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
        .Formula = "=CONCATENATE(""" & Cells((i + 1), 1) & ""","" "",""" & Cells((i + 1), 2) & """)"
        .Font.Bold = True
    End With
    
    ' Initialize the value of k
    k = 1
    
    ' Find and delete the redundant rows
    While k <= Cells(Rows.Count, 10 + (i - 1) * (ccount_01 + 1)).End(xlUp).Row - 2
        If (ActiveSheet.ListObjects(tname_02).DataBodyRange(k, 2) = "" And ActiveSheet.ListObjects(tname_02).DataBodyRange(k, 3) = "" And _
        ActiveSheet.ListObjects(tname_02).DataBodyRange(k, 4) = "" And ActiveSheet.ListObjects(tname_02).DataBodyRange(k, 5) = "") Or _
        (ActiveSheet.ListObjects(tname_02).DataBodyRange(k, 2) = 0 And ActiveSheet.ListObjects(tname_02).DataBodyRange(k, 3) = 0 And _
        ActiveSheet.ListObjects(tname_02).DataBodyRange(k, 4) = 0 And ActiveSheet.ListObjects(tname_02).DataBodyRange(k, 5) = 0) Then
            ActiveSheet.ListObjects(tname_02).Range.Select
            Selection.ListObject.ListRows(k).Delete
        End If
        k = k + 1
    Wend
    
    ' Add a stock chart
    Range(tname_02 & "[[#All],[日期]:[收市]]").Select
    ActiveSheet.Shapes.AddChart2(322, xlStockOHLC).Select
    With Selection
        .Height = Range(Cells(Cells(Rows.Count, 10 + (i - 1) * (ccount_01 + 1)).End(xlUp).Row + 2, 10 + (i - 1) * (ccount_01 + 1)), _
        Cells(Cells(Rows.Count, 10 + (i - 1) * (ccount_01 + 1)).End(xlUp).Row + 21, 10 + (i - 1) * (ccount_01 + 1))).Height - 5
        .Width = Range(Cells(Cells(Rows.Count, 10 + (i - 1) * (ccount_01 + 1)).End(xlUp).Row + 2, 10 + (i - 1) * (ccount_01 + 1)), _
        Cells(Cells(Rows.Count, 10 + (i - 1) * (ccount_01 + 1)).End(xlUp).Row + 2, 10 + (i - 1) * (ccount_01 + 1) + 6)).Width - 5
        .Top = Cells(Cells(Rows.Count, 10 + (i - 1) * (ccount_01 + 1)).End(xlUp).Row + 2, 10 + (i - 1) * (ccount_01 + 1)).Top
        .Left = Cells(Cells(Rows.Count, 10 + (i - 1) * (ccount_01 + 1)).End(xlUp).Row + 2, 10 + (i - 1) * (ccount_01 + 1)).Left
    End With
    
    ' Edit the chart
    ActiveChart.Parent.Name = ("Chart_" & i)
    ActiveChart.ChartTitle.Select
    Selection.Caption = "=主頁面!R1C" & (10 + (i - 1) * (ccount_01 + 1))
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = _
    WorksheetFunction.RoundDown( _
    WorksheetFunction.Min(Range(tname_02 & "[[#All],[開市]:[收市]]")), _
    2 - (1 + Int(WorksheetFunction.Log10(Abs(WorksheetFunction.Min(Range(tname_02 & "[[#All],[開市]:[收市]]")))))))
    ActiveChart.Axes(xlValue).MaximumScale = _
    WorksheetFunction.RoundUp( _
    WorksheetFunction.Max(Range(tname_02 & "[[#All],[開市]:[收市]]")), _
    2 - (1 + Int(WorksheetFunction.Log10(Abs(WorksheetFunction.Min(Range(tname_02 & "[[#All],[開市]:[收市]]")))))))
    ActiveChart.Legend.Select
    Selection.Delete
    
' End the For loop
Next i
End Sub
