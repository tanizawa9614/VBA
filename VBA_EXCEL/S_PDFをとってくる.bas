Attribute VB_Name = "S_PDF���Ƃ��Ă���"
Option Explicit

Sub PDF���Ƃ��Ă���()  '�w��t�H���_��PDF��V�K�t�H���_��
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
      For Each fl In FSO.GetFolder(fpath).Files
         ext = FSO.GetExtensionName(fl.Path)
         If InStr(ext, "pdf") > 0 Then
            FSO.CopyFile fl.Path, "C:\Users\yuuki\OneDrive - Osaka University\�f�X�N�g�b�v\����w_�@���ߋ���_PDF��\"
         End If
      Next
   Next
   Application.ScreenUpdating = True
End Sub
