Attribute VB_Name = "Module1"
Sub data_enquiry_01()
'
' data_enquiry_01 ����
'

' Activate the target worksheet
Worksheets("�W�����q�C��").Activate

' Delete the original table
Columns(1).Resize(, 2).ClearContents
Columns(1).Resize(, 2).ClearFormats

' Clear the original query
For Each qr In ThisWorkbook.Queries
    If qr = "�W�����q�C��" Then
        qr.Delete
    End If
Next qr

' Clear the existing connections
For Each cn In ThisWorkbook.Connections
    cn.Delete
Next cn

    ' Import the data
    ActiveWorkbook.Queries.Add Name:="�W�����q�C��", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "�ӷ� = Web.Page(Web.Contents(""http://www.hkexnews.hk/hyperlink/hyperlist_c.HTM""))," & _
        Chr(13) & "" & Chr(10) & "    Data5 = �ӷ�{5}[Data]," & Chr(13) & "" & Chr(10) & _
        "�w�ܧ����� = Table.TransformColumnTypes(Data5,{{""�Ѳ��N��"", type text}, {""�W�����q���W��"", type text}, {""�W�����q�����}"", type text}})," & _
        Chr(13) & "" & Chr(10) & "�w������Ʀ� = Table.RemoveColumns(�w�ܧ�����,{""�W�����q�����}""})" & _
        Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "�w������Ʀ�"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""�W�����q�C��""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [�W�����q�C��]")
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
        .ListObject.DisplayName = "�W�����q�C��"
        .Refresh BackgroundQuery:=False
    End With
    
    ' Tweaking
    ActiveSheet.ListObjects("�W�����q�C��").TableStyle = "TableStyleMedium13"
End Sub
