Attribute VB_Name = "S_�摜��PDF�ɕϊ�"
Option Explicit

Sub �摜��PDF�ɕϊ�()
   Dim fpath As String
   Dim i As Double
   Dim FSO As Object, fl As Object, ext As String
   Set FSO = CreateObject("Scripting.FileSystemObject")
   Dim buf As String, doc As Document
   Dim NewPDFName As String
   
   Application.EnableCancelKey = wdCancelInterrupt
   Do
      Application.ScreenUpdating = False
L1:
      With Application.FileDialog(msoFileDialogFolderPicker)
         If .Show = True Then fpath = .SelectedItems(1)
         If fpath = "" Then GoTo L1
      End With
      
      NewPDFName = fpath & "\" & FSO.GetFolder(fpath).Name & "_pdf��.pdf"
      If FSO.FileExists(NewPDFName) Then
         buf = MsgBox("����" & NewPDFName & "�Ƃ����t�@�C�������݂��܂�" & vbCr _
         & "���̃t�H���_��I�����Ă�������" & vbCr & "�������I�����܂����H", vbYesNo)
         If buf = vbNo Then GoTo L1
         Exit Sub
      End If
      
      Set doc = Documents.Add(DocumentType:=wdNewBlankDocument)
      
      i = 0
      With doc.PageSetup
         .TopMargin = MillimetersToPoints(i)
         .BottomMargin = MillimetersToPoints(i)
         .LeftMargin = MillimetersToPoints(i)
         .RightMargin = MillimetersToPoints(i)
      End With
      
      For Each fl In FSO.GetFolder(fpath).Files
         ext = FSO.GetExtensionName(fl.path)
         Select Case ext
         Case "jpg", "JPG", "png", "PNG", "jpeg", "JPEG", "jpe", "jfif", "pjpeg", "pjp"
            doc.Bookmarks("\EndOfDoc").Range _
             .InlineShapes.AddPicture FileName:=fl.path
         End Select
      Next
      
      doc.ExportAsFixedFormat _
      OutputFileName:=NewPDFName, _
      ExportFormat:=wdExportFormatPDF
      
      doc.Close saveChanges:=False
      
      Application.ScreenUpdating = True
      buf = MsgBox("�������I�����܂����H", vbQuestion + vbYesNo)
      If buf = vbYes Then
         Shell "C:\Windows\Explorer.exe " & fpath, vbNormalFocus
         GoTo L2
      End If
   Loop
L2:
End Sub
