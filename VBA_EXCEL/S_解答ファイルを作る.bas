Attribute VB_Name = "S_�𓚃t�@�C�������"
Option Explicit

Sub �𓚃t�H���_�����()
   Dim path As String, ext As String
   Dim FSO As Object, fl As Object, NewName As String
   Const C As String = "��"
   
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

