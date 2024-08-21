Attribute VB_Name = "S_�摜PDF�ꊇ_1���w"
Option Explicit

Sub �摜PDF�ꊇ()  '�w��t�H���_�̉��w�t�H���_���ꊇ�ύX
   Dim fpath As String, pfpath As String
   Dim i As Double, fol As Object, sfol As Object
   Dim FSO As Object, fl As Object, ext As String
   Set FSO = CreateObject("Scripting.FileSystemObject")
   Dim buf As String, doc As Document
   Dim NewPDFName As String
   
   With Application.FileDialog(msoFileDialogFolderPicker)
         If .Show = True Then pfpath = .SelectedItems(1)
   End With
   
   Application.ScreenUpdating = True
   
   For Each fol In FSO.GetFolder(pfpath).SubFolders
      For Each sfol In FSO.GetFolder(fol).SubFolders
         fpath = sfol.path
         Set doc = Documents.Add(DocumentType:=wdNewBlankDocument)
         NewPDFName = fpath & "\" & FSO.GetFolder(fpath).Name & "_pdf��.pdf"
         i = 0
         With doc.PageSetup
            .TopMargin = MillimetersToPoints(i)
            .BottomMargin = MillimetersToPoints(i)
            .LeftMargin = MillimetersToPoints(i)
            .RightMargin = MillimetersToPoints(i)
         End With
            
         For Each fl In FSO.GetFolder(fpath).Files
            ext = FSO.GetExtensionName(fl.path)
            If InStr(ext, "jpg") > 0 Then
               buf = TempName(fl.Name)
               If fl.Name <> buf Then
                  fl.Name = buf
               End If
            End If
         Next
            
         For Each fl In FSO.GetFolder(fpath).Files
            ext = FSO.GetExtensionName(fl.path)
            If InStr(ext, "jpg") > 0 Then
               With doc.Bookmarks("\EndOfDoc").Range
                  .InlineShapes.AddPicture FileName:=fl.path
               End With
            End If
         Next
            
         For Each fl In FSO.GetFolder(fpath).Files
            ext = FSO.GetExtensionName(fl.path)
            If InStr(ext, "jpg") > 0 Then
               buf = TempNameDelete(fl.Name)
               If fl.Name <> buf Then
                  fl.Name = buf
               End If
            End If
         Next
            
         doc.ExportAsFixedFormat _
         OutputFileName:=NewPDFName, _
         ExportFormat:=wdExportFormatPDF
            
         doc.Close saveChanges:=False
      Next
   Next
   Application.ScreenUpdating = True
End Sub
Function TempName(A As String) As String
   Dim B(5), i As Long
   B(0) = "���w"
   B(1) = "���"
   B(2) = "�\��"
   B(3) = "����"
   B(4) = "�y��"
   B(5) = "�v��"
   For i = 0 To 5
      If InStr(A, B(i)) > 0 Then
         TempName = i & A
         Exit Function
      End If
   Next
   TempName = A
End Function
Function TempNameDelete(A As String) As String
   Dim n As String
   n = Left(A, 1)
   If IsNumeric(n) Then
      TempNameDelete = Mid(A, 2, Len(A))
   Else
      TempNameDelete = A
   End If
End Function

