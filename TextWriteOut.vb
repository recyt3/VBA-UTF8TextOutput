Sub GermanTxt()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    Dim datFile As String
    datFile = ActiveWorkbook.Path & "\data.txt"
    Open datFile For Output As #1
    Dim i As Long
    i = 1
    Do While ws.Cells(i, 6).Value <> ""
        Print #1, ws.Cells(i, 6).Value
        i = i + 1
    Loop
    Close #1
    MsgBox "data.txtに書き出しました"
End Sub

