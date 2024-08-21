Attribute VB_Name = "S_任意フォルダ立ち上げ操作"
Option Explicit

Sub 任意フォルダ立ち上げ操作()
   Dim Path As String
   Dim FSO As Object, fl As Object
   Set FSO = CreateObject("Scripting.FileSystemObject")
   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then Path = .SelectedItems(1)
   End With
   For Each fl In FSO.GetFolder(Path).Files
'      fl.Name = Replace(fl.Name, "_pdf版", "")
   Next
End Sub
