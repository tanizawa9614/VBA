Attribute VB_Name = "S_フォルダ全展開"
Option Explicit

Sub Main1()
   Call 親フォルダ上に全展開_親フォルダ指定_複数フォルダ
End Sub

Sub Main2()
   Call 親フォルダ上に全展開_サブフォルダ指定_1つのフォルダのみ
End Sub

Sub 親フォルダ上に全展開_親フォルダ指定_複数フォルダ()
   Dim path As String
   Dim FSO As Object, fl As Object, sfl As Object
   Set FSO = CreateObject("Scripting.FileSystemObject")

   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then path = .SelectedItems(1)
   End With
   
   For Each sfl In FSO.GetFolder(path).SubFolders
      For Each fl In FSO.GetFolder(sfl.path).Files
         FSO.GetFile(fl.path).Move path & "\"
      Next fl
   Next sfl
   
   Call 空フォルダの削除(path)
   
   MsgBox "終了しました"
   
   Set FSO = Nothing
   Set fl = Nothing
   Set sfl = Nothing
End Sub

Sub 親フォルダ上に全展開_サブフォルダ指定_1つのフォルダのみ()
   Dim path As String
   Dim FSO As Object, fl As Object, p_fol As String
   Set FSO = CreateObject("Scripting.FileSystemObject")
   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then path = .SelectedItems(1)
   End With
   p_fol = FSO.GetFolder(path).ParentFolder
   For Each fl In FSO.GetFolder(path).Files
      FSO.GetFile(fl.path).Move p_fol & "\"
   Next fl
   
   Call 空フォルダの削除(path)
   
   MsgBox "終了しました"
   
   Set FSO = Nothing
   Set fl = Nothing
End Sub

Sub 空フォルダの削除(path As String)
   Dim flag As String, sfl As Object
   Dim FSO As Object
   Set FSO = CreateObject("Scripting.FileSystemObject")
   
   flag = MsgBox("空フォルダを削除しますか？", vbYesNo)
   If flag = vbYes Then
      For Each sfl In FSO.GetFolder(path).SubFolders
         If FSO.GetFolder(sfl.path).SubFolders.Count >= 1 Then
            MsgBox "場所：" & sfl.path & vbCr _
            & "フォルダ名：" & sfl.Name & vbCr & "にはフォルダが存在します"
            GoTo L1
         End If
         If FSO.GetFolder(sfl.path).Files.Count >= 1 Then
            MsgBox "場所：" & sfl.path & vbCr _
            & "フォルダ名：" & sfl.Name & vbCr & "にはファイルが存在します"
            GoTo L1
         End If
         FSO.DeleteFolder sfl.path
L1:
      Next
   End If
End Sub


