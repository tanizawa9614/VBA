Attribute VB_Name = "S_�t�H���_�S�W�J"
Option Explicit

Sub Main1()
   Call �e�t�H���_��ɑS�W�J_�e�t�H���_�w��_�����t�H���_
End Sub

Sub Main2()
   Call �e�t�H���_��ɑS�W�J_�T�u�t�H���_�w��_1�̃t�H���_�̂�
End Sub

Sub �e�t�H���_��ɑS�W�J_�e�t�H���_�w��_�����t�H���_()
   Dim path As String
   Dim FSO As Object, fl As Object, sfl As Object
   Set FSO = CreateObject("Scripting.FileSystemObject")

   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then path = .SelectedItems(1)
   End With
   
   For Each sfl In FSO.GetFolder(path).SubFolders
      For Each fl In FSO.GetFolder(sfl.path).Files
         FSO.GetFile(fl.path).Move path & "\"
      Next fl
   Next sfl
   
   Call ��t�H���_�̍폜(path)
   
   MsgBox "�I�����܂���"
   
   Set FSO = Nothing
   Set fl = Nothing
   Set sfl = Nothing
End Sub

Sub �e�t�H���_��ɑS�W�J_�T�u�t�H���_�w��_1�̃t�H���_�̂�()
   Dim path As String
   Dim FSO As Object, fl As Object, p_fol As String
   Set FSO = CreateObject("Scripting.FileSystemObject")
   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then path = .SelectedItems(1)
   End With
   p_fol = FSO.GetFolder(path).ParentFolder
   For Each fl In FSO.GetFolder(path).Files
      FSO.GetFile(fl.path).Move p_fol & "\"
   Next fl
   
   Call ��t�H���_�̍폜(path)
   
   MsgBox "�I�����܂���"
   
   Set FSO = Nothing
   Set fl = Nothing
End Sub

Sub ��t�H���_�̍폜(path As String)
   Dim flag As String, sfl As Object
   Dim FSO As Object
   Set FSO = CreateObject("Scripting.FileSystemObject")
   
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


