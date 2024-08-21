Attribute VB_Name = "S_院試用_解答名前"
Option Explicit

Sub 院試用_解答フォルダ名前変更()  '指定フォルダのPDFを新規フォルダに
   Dim fpath As String, pfpath As String
   Dim i As Double, fol As Object
   Dim FSO As Object, fl As Object, ext As String
   Set FSO = CreateObject("Scripting.FileSystemObject")
   Dim buf As String
   Dim NewPDFName As String
   
   With Application.FileDialog(msoFileDialogFolderPicker)
         If .Show = True Then pfpath = .SelectedItems(1)
   End With
   
   Application.ScreenUpdating = True
   
   For Each fol In FSO.GetFolder(pfpath).SubFolders
      fpath = fol.Path
      NewPDFName = fpath & "\" & FSO.GetFolder(fpath).Name & "_pdf版.pdf"
                  
      For Each fl In FSO.GetFolder(fpath).SubFolders
         If InStr(fl.Name, "解答") > 0 Then
            fl.Name = fol.Name & fl.Name
         End If
      Next
   Next
   Application.ScreenUpdating = True
End Sub
