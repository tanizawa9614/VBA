Attribute VB_Name = "S_動画再生時間"
Option Explicit

Sub 動画再生時間()
   Dim Path As String, i As Long, t()
   Dim j As Long
   Dim FSO As Object, fl As Object
   Set FSO = CreateObject("Scripting.FileSystemObject")
   
   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then Path = .SelectedItems(1)
   End With
   
   Dim Shel As Object, Foldr As Object
   Set Shel = CreateObject("Shell.Application")
   Dim shFolder As Object
   Set shFolder = Shel.Namespace(Path & "\")
   
   ReDim t(FSO.GetFolder(Path).Files.Count, 500)
   For Each fl In FSO.GetFolder(Path).Files
'      Cells(i + 1, 1) = shFolder.GetDetailsOf(shFolder.ParseName(fl.Path), 0)
      For j = 0 To 500
         t(i, j) = shFolder.GetDetailsOf _
            (shFolder.ParseName(fl.Path), j)
      Next
      i = i + 1
   Next fl
   
   
   MsgBox "終了しました"
   
   Set FSO = Nothing
   Set fl = Nothing
   Set Shel = Nothing
   Set shFolder = Nothing
End Sub
