Attribute VB_Name = "Sub_空フォルダの削除"
Option Explicit

Sub 空フォルダの削除()
   Dim path As String, buf As String
   Dim FSO As Object, fl As Object, sfl As Object
   Set FSO = CreateObject("Scripting.FileSystemObject")

   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then path = .SelectedItems(1)
   End With
   
   For Each sfl In FSO.GetFolder(path).SubFolders
      If FSO.GetFolder(sfl.path).SubFolders.Count >= 1 Then
'        MsgBox "場所：" & sfl.path & vbCr _
         & "フォルダ名：" & sfl.Name & vbCr & "にはフォルダが存在します"
         GoTo L1
      End If
      If FSO.GetFolder(sfl.path).Files.Count >= 1 Then
'        MsgBox "場所：" & sfl.path & vbCr _
         & "フォルダ名：" & sfl.Name & vbCr & "にはファイルが存在します"
         GoTo L1
      End If
      buf = MsgBox("場所：" & sfl.path & vbCr _
         & "フォルダ名：" & sfl.Name & vbCr & "を削除しますか？" _
         , vbYesNo)
      If buf = vbYes Then FSO.DeleteFolder sfl.path
L1:
   Next
   MsgBox "終了しました"
End Sub


