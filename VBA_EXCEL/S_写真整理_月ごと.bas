Attribute VB_Name = "S_�ʐ^����_������"
Option Explicit
Sub Main1()
   Call �ʐ^����_������
End Sub
Sub Main2()
   Call �e�t�H���_��ɑS�W�J_�e�t�H���_�w��_�����t�H���_
End Sub
Sub Main3()
   Call ��t�H���_�̍폜
End Sub
Sub �ʐ^����_������()
    Dim place As String, FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim fl As Object
    Dim New_Folder As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then place = .SelectedItems(1)
        If place = "" Then Exit Sub
    End With
    
'    On Error Resume Next
    Dim tmpfname As String
    For Each fl In FSO.GetFolder(place).Files
        DoEvents
        '�t�@�C���́u�쐬���v���擾,�V�K�t�H���_���́u���t�v
        New_Folder = Format(FileDateTime(fl.path), "yyyym")
        New_Folder = Left(New_Folder, 4) & "�N" & Mid(New_Folder, 5) & "��"
        
        If Not FSO.FolderExists(place & "\" & New_Folder) Then
            FSO.CreateFolder (place & "\" & New_Folder)
            '�V�K�t�H���_���쐬
        End If
        '�t�H���_�́u�쐬���v�̃t�H���_�Ɉړ�
        tmpfname = fl.Name
        Do While FSO.FileExists(place & "\" & New_Folder & "\" & tmpfname)
            tmpfname = "1" & tmpfname
        Loop
        If fl.Name <> tmpfname Then fl.Name = tmpfname
        FSO.MoveFile fl.path, place & "\" & New_Folder & "\"
    Next
    MsgBox "�I�����܂���"
    Set FSO = Nothing
    Set fl = Nothing
End Sub

Sub �e�t�H���_��ɑS�W�J_�e�t�H���_�w��_�����t�H���_()
    Dim A As String, B As String
    Dim place As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then place = .SelectedItems(1)
        If place = "" Then Exit Sub
    End With
    
    Dim pfl As Object
    Set pfl = FSO.GetFolder(place)
    
    Dim subfl As Object
    Dim fl As Object, tmpname As String
    For Each subfl In pfl.SubFolders
        DoEvents
        If subfl.Files.Count >= 1 Then
            For Each fl In subfl.Files
                tmpname = fl.Name
                Do While FSO.FileExists(pfl.path & "\" & tmpname)
                    tmpname = "1" & tmpname
                Loop
                If fl.Name <> tmpname Then fl.Name = tmpname
                FSO.GetFile(fl.path).Move pfl.path & "\"
            Next
        End If
    Next
    
    Call delete_emptyfile(place)
    
    MsgBox "�I�����܂���"
    
    Set FSO = Nothing
    Set fl = Nothing
    
End Sub
Sub ��t�H���_�̍폜()
    Dim FolPath As String, i As Long
    Dim FSO As Object, fl As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then FolPath = .SelectedItems(1)
        If FolPath = "" Then Exit Sub
    End With
    For Each fl In FSO.GetFolder(FolPath).SubFolders
        If fl.Files.Count = 0 And fl.SubFolders.Count = 0 Then
'            MsgBox fl.Name
            fl.Delete
            DoEvents
        End If
    Next
    MsgBox "�I�����܂���"
End Sub
Private Sub delete_emptyfile(Optional path As String)
    Dim flag As String, sfl As Object
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If path = "" Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = True Then path = .SelectedItems(1)
            If path = "" Then Exit Sub
        End With
    End If
    flag = MsgBox("��t�H���_���폜���܂����H", vbYesNo)
    If flag = vbYes Then
    For Each sfl In FSO.GetFolder(path).SubFolders
        If FSO.GetFolder(sfl.path).SubFolders.Count >= 1 Then
            MsgBox "�ꏊ�F" & sfl.path & vbCr _
            & "�t�H���_���F" & sfl.Name & vbCr & "�ɂ̓t�H���_�����݂��܂�"
            GoTo L1
        End If
        If FSO.GetFolder(sfl.path).Files.Count >= 1 Then
            MsgBox "�ꏊ�F" & sfl.path & vbCr _
            & "�t�H���_���F" & sfl.Name & vbCr & "�ɂ̓t�@�C�������݂��܂�"
            GoTo L1
        End If
        FSO.DeleteFolder sfl.path
L1:
    Next
    End If
End Sub
