Attribute VB_Name = "S_FileDrop"
Option Explicit

Sub �t�@�C�����h���b�v���Ď擾()
    FileDropUF.Show
End Sub

Sub �V�[�g�̑S�폜()
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
