Attribute VB_Name = "Module1"
Sub data_enquiry_01()
'
' data_enquiry_01 巨集
'

' Activate the target worksheet
Worksheets("上市公司列表").Activate

' Delete the original table
Columns(1).Resize(, 2).ClearContents
Columns(1).Resize(, 2).ClearFormats

' Clear the original query
For Each qr In ThisWorkbook.Queries
    If qr = "上市公司列表" Then
        qr.Delete
    End If
Next qr

' Clear the existing connections
For Each cn In ThisWorkbook.Connections
    cn.Delete
Next cn

    ' Import the data
    ActiveWorkbook.Queries.Add Name:="上市公司列表", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "來源 = Web.Page(Web.Contents(""http://www.hkexnews.hk/hyperlink/hyperlist_c.HTM""))," & _
        Chr(13) & "" & Chr(10) & "    Data5 = 來源{5}[Data]," & Chr(13) & "" & Chr(10) & _
        "已變更類型 = Table.TransformColumnTypes(Data5,{{""股票代號"", type text}, {""上市公司之名稱"", type text}, {""上市公司之網址"", type text}})," & _
        Chr(13) & "" & Chr(10) & "已移除資料行 = Table.RemoveColumns(已變更類型,{""上市公司之網址""})" & _
        Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "已移除資料行"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""上市公司列表""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [上市公司列表]")
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
        .ListObject.DisplayName = "上市公司列表"
        .Refresh BackgroundQuery:=False
    End With
    
    ' Tweaking
    ActiveSheet.ListObjects("上市公司列表").TableStyle = "TableStyleMedium13"
End Sub
