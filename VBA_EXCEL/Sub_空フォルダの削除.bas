Attribute VB_Name = "Sub_��t�H���_�̍폜"
Option Explicit

Sub ��t�H���_�̍폜()
   Dim path As String, buf As String
   Dim FSO As Object, fl As Object, sfl As Object
   Set FSO = CreateObject("Scripting.FileSystemObject")

   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then path = .SelectedItems(1)
   End With
   
   For Each sfl In FSO.GetFolder(path).SubFolders
      If FSO.GetFolder(sfl.path).SubFolders.Count >= 1 Then
'        MsgBox "�ꏊ�F" & sfl.path & vbCr _
         & "�t�H���_���F" & sfl.Name & vbCr & "�ɂ̓t�H���_�����݂��܂�"
         GoTo L1
      End If
      If FSO.GetFolder(sfl.path).Files.Count >= 1 Then
'        MsgBox "�ꏊ�F" & sfl.path & vbCr _
         & "�t�H���_���F" & sfl.Name & vbCr & "�ɂ̓t�@�C�������݂��܂�"
         GoTo L1
      End If
      buf = MsgBox("�ꏊ�F" & sfl.path & vbCr _
         & "�t�H���_���F" & sfl.Name & vbCr & "���폜���܂����H" _
         , vbYesNo)
      If buf = vbYes Then FSO.DeleteFolder sfl.path
L1:
   Next
   MsgBox "�I�����܂���"
End Sub


