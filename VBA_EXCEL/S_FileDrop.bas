Attribute VB_Name = "S_FileDrop"
Option Explicit

Sub ファイルをドロップして取得()
    FileDropUF.Show
End Sub

Sub シートの全削除()
    Dim n As Long
    Dim i As Long
    Application.DisplayAlerts = False
    With ThisWorkbook
        n = .Sheets.Count
        For i = 2 To n
            .Sheets(2).Delete
        Next
    End With
    Application.DisplayAlerts = True
End Sub
