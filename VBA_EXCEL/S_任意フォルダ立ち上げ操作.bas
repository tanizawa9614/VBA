Attribute VB_Name = "S_�C�Ӄt�H���_�����グ����"
Option Explicit

Sub �C�Ӄt�H���_�����グ����()
   Dim Path As String
   Dim FSO As Object, fl As Object
   Set FSO = CreateObject("Scripting.FileSystemObject")
   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then Path = .SelectedItems(1)
   End With
   For Each fl In FSO.GetFolder(Path).Files
'      fl.Name = Replace(fl.Name, "_pdf��", "")
   Next
End Sub
