Attribute VB_Name = "S_ファイル操作ひな形"
Option Explicit

Sub ファイル操作ひな形()
   Dim FolPath As String
   Dim FSO As Object, fl As Object
   Set FSO = CreateObject("Scripting.FileSystemObject")
   
   With Application.FileDialog(msoFileDialogFolderPicker)
       If .Show = True Then FolPath = .SelectedItems(1)
   End With

   For Each fl In FSO.GetFolder(FolPath).Files
      Cells(i + 1, 1) = fl.Name
      i = i + 1
   Next
End Sub
