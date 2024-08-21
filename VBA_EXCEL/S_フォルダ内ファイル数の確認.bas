Attribute VB_Name = "S_フォルダ内ファイル数の確認"
Option Explicit

Sub フォルダ内ファイル数の確認()
   Dim path As String, i As Long
   Dim FSO As Object, sfl As Object
   Set FSO = CreateObject("Scripting.FileSystemObject")

   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then path = .SelectedItems(1)
   End With
   
   Cells(1, 1) = "名前"
   Cells(1, 2) = "フォルダ数"
   
   For Each sfl In FSO.GetFolder(path).SubFolders
      Cells(i + 2, 1) = sfl.Name
      Cells(i + 2, 2) = sfl.Files.Count
      i = i + 1
   Next sfl
   
   
   MsgBox "終了しました"
   
   Set FSO = Nothing
   Set sfl = Nothing
End Sub
