Attribute VB_Name = "S_�t�H���_���t�@�C�����̊m�F"
Option Explicit

Sub �t�H���_���t�@�C�����̊m�F()
   Dim path As String, i As Long
   Dim FSO As Object, sfl As Object
   Set FSO = CreateObject("Scripting.FileSystemObject")

   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then path = .SelectedItems(1)
   End With
   
   Cells(1, 1) = "���O"
   Cells(1, 2) = "�t�H���_��"
   
   For Each sfl In FSO.GetFolder(path).SubFolders
      Cells(i + 2, 1) = sfl.Name
      Cells(i + 2, 2) = sfl.Files.Count
      i = i + 1
   Next sfl
   
   
   MsgBox "�I�����܂���"
   
   Set FSO = Nothing
   Set sfl = Nothing
End Sub
