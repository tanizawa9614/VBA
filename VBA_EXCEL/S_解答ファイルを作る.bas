Attribute VB_Name = "S_解答ファイルを作る"
Option Explicit

Sub 解答フォルダを作る()
   Dim path As String, ext As String
   Dim FSO As Object, fl As Object, NewName As String
   Const C As String = "解答"
   
   Set FSO = CreateObject("Scripting.FileSystemObject")
   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then path = .SelectedItems(1)
   End With
   
   For Each fl In FSO.GetFolder(path).Files
      ext = "." & FSO.GetExtensionName(fl.path)
      If FSO.FolderExists(path & "\" & C) = False _
      Then FSO.CreateFolder path & "\" & C
      If InStr(fl.Name, C) > 0 Then FSO.MoveFile fl.path, path & "\" & C & "\"
   Next fl
   Set FSO = Nothing
End Sub

