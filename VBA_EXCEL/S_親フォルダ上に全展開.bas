Attribute VB_Name = "S_親フォルダ上に全展開"
Option Explicit

Sub 親フォルダ上に全展開()
   Dim path As String, ext As String
   Dim FSO As Object, fl As Object, p_fol As String
   Set FSO = CreateObject("Scripting.FileSystemObject")
   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then path = .SelectedItems(1)
   End With
   p_fol = FSO.GetFolder(path).ParentFolder
   For Each fl In FSO.GetFolder(path).Files
      MsgBox fl.path
      FSO.GetFile(fl.path).Move p_fol & "\"
   Next fl
End Sub
 
