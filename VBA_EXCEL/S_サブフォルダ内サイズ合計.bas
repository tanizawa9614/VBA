Attribute VB_Name = "S_�T�u�t�H���_���T�C�Y���v"
Option Explicit

Sub �T�u�t�H���_���̃t�@�C���T�C�Y���v()
   Dim Path As String, fl As Object
   Dim FSO As Object, sfl As Object
   Dim F_Size As Double
   Set FSO = CreateObject("Scripting.FileSystemObject")
   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then Path = .SelectedItems(1)
   End With
   For Each sfl In FSO.GetFolder(Path).SubFolders
      For Each fl In FSO.GetFolder(sfl).Files
         F_Size = F_Size + fl.Size / 1024 ^ 2
      Next
   Next
   MsgBox F_Size
End Sub
