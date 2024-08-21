Attribute VB_Name = "S_PDF�ϊ�"
Option Explicit

Sub PDF�ɕϊ�()
    Dim FolPath As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then FolPath = .SelectedItems(1)
    End With
    
    If Len(FolPath) = 0 Then Exit Sub
    
    If Not FSO.FolderExists(FolPath & "\PDF\") Then '�t�H���_���݊m�F
        MkDir FolPath & "\PDF\" '�t�H���_�쐬
    End If
    
    Dim fl As Object
    Application.ScreenUpdating = False
    For Each fl In FSO.GetFolder(FolPath).Files
        If ThisWorkbook.FullName <> fl.Path Then
            Call Conv_PDF(fl.ParentFolder.Path, Left(fl.Name, InStrRev(fl.Name, ".") - 1), FSO.GetExtensionName(fl.Path))
        End If
    Next
    Set fl = Nothing
    Application.ScreenUpdating = True
    
    MsgBox "�I�����܂���"
    FolPath = FolPath & "\PDF\"
    Shell "C:\Windows\Explorer.exe " & FolPath, vbNormalFocus
End Sub

Private Sub Conv_PDF(P_Path As String, fl_Name As String, ext As String)
    Dim NewPath As String
    Dim objOffice As Object
    
    NewPath = P_Path & "\PDF\" & fl_Name & ".pdf"
    P_Path = P_Path & "\" & fl_Name & "." & ext
    
    Select Case ext
        Case "xls", "xlsx" 'Excel97-2003,Excel2007�ȍ~
            Set objOffice = Excel.Application
            With objOffice.Workbooks.Open(P_Path)
                .ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=NewPath, Openafterpublish:=False
                .Close
            End With
        Case "doc", "docx" 'Word97-2003,Word2007�ȍ~
            Set objOffice = CreateObject("Word.Application")
            With objOffice.Documents.Open(P_Path)
                .ExportAsFixedFormat OutputFileName:=NewPath, _
                ExportFormat:=17
                .Close
            End With
        Case "ppt", "pptx" 'Powerpoint97-2003,Powerpoint2007�ȍ~
            Set objOffice = CreateObject("Powerpoint.Application")
            With objOffice.Presentations.Open(P_Path)
                .SaveAs Filename:=NewPath, FileFormat:=32
                .Close
            End With
    End Select
    Set objOffice = Nothing
End Sub

