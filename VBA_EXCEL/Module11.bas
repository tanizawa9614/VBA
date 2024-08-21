Attribute VB_Name = "Module11"
Option Explicit

Sub Convert_to_PDF()
    Dim strDirPath As String
    strDirPath = Search_Directory() '�t�H���_�̑I��
    If Len(strDirPath) = 0 Then Exit Sub
    Call Make_Dir(strDirPath, "\PDF") '�t�H���_�쐬
    Call Search_Files(strDirPath)
    MsgBox "�I�����܂���"
End Sub

Private Function Search_Directory() As String '�t�H���_�̑I��
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then Search_Directory = .SelectedItems(1)
    End With
End Function

Private Sub Make_Dir(ByVal Path As String, ByVal Dn As String)
    If Dir(Path & Dn, vbDirectory) = "" Then '�t�H���_���݊m�F
        MkDir Path & Dn '�t�H���_�쐬
    End If
End Sub

Private Sub Search_Files(ByVal Path As String)
    Dim strFile As String
    strFile = Dir(Path & "\" & "*.*") '�t�@�C���m�F
    Application.ScreenUpdating = False
    Do Until strFile = ""
        If ThisWorkbook.FullName <> Path & "\" & strFile Then
            Call Conv_PDF(Path, "\" & strFile)
        End If
        strFile = Dir() '���̃t�@�C���m�F
    Loop
    Application.ScreenUpdating = True
End Sub

Private Function Get_Extension(ByVal Path As String) As String '�g���q�擾
    Dim i As Long
    i = InStrRev(Path, ".", -1, vbTextCompare)
    If i = 0 Then Exit Function
    Get_Extension = Mid$(Path, i + 1)
End Function

Private Sub Conv_PDF(ByVal Path As String, ByVal Fn As String)
    Dim filePath  As String
    Dim objOffice As Object
    filePath = Path & "\PDF" & Left$(Fn, InStrRev(Fn, ".")) & "pdf"
    Path = Path & Fn
    Select Case Get_Extension(Fn) '�t�@�C��������g���q�擾
        Case "xls", "xlsx" 'Excel97-2003,Excel2007�ȍ~
            Set objOffice = Excel.Application
            With objOffice.Workbooks.Open(Path)
                .ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=filePath, Openafterpublish:=False
                .Close
            End With
        Case "doc", "docx" 'Word97-2003,Word2007�ȍ~
            Set objOffice = CreateObject("Word.Application")
            With objOffice.Documents.Open(Path)
                .ExportAsFixedFormat OutputFileName:=filePath, _
                ExportFormat:=17
                .Close
            End With
            objOffice.Quit
        Case "ppt", "pptx" 'Powerpoint97-2003,Powerpoint2007�ȍ~
            Set objOffice = CreateObject("Powerpoint.Application")
            With objOffice.Presentations.Open(Path)
                .SaveAs Filename:=filePath, FileFormat:=32
                .Close
            End With
            objOffice.Quit
    End Select
End Sub

