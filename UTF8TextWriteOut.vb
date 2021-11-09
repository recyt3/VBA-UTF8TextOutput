Function FileOutput(ByVal filepathName As String, ByVal firstCol As Integer, ByVal firstRow As Integer) As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    Dim datFile As String
    
    ' ファイルチェックと削除
    If Dir(filepathName) <> "" Then
        Kill (filepathName)
    End If
    'ADODB.Streamオブジェクトを生成
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
    With adoSt
        .Mode = 3
        .Type = 2  'オブジェクトに保存するデータの種類を文字列型に指定する
        .Charset = "UTF-8"
        .Open
        
        i = firstCol
        Do While ws.Cells(i, firstRow).Value <> ""
            .WriteText ws.Cells(i, firstRow).Value & vbLf
            i = i + 1
        Loop
        .SaveToFile (filepathName)
    End With

    MsgBox filepathName + "に書き出しました"
    FileOutput = 0
End Function
