Attribute VB_Name = "F_PDFCONV"
Option Explicit
Function PDFCONV(�Ώ� As Range) As String
   Dim PDF_path As String
   Dim path As Range, i As Long
   Dim Office As Object
   Dim flag As String
   flag = MsgBox("PDF�ւ̕ϊ����J�n���܂����H", vbQuestion + vbYesNo)
   If flag = vbNo Then
      PDFCONV = "�����𒆎~���܂���"
      Exit Function
   End If
   For Each path In �Ώ�
      PDF_path = Left(path, InStrRev(path, ".")) & "pdf"
      Select Case Mid(path, InStrRev(path, ".") + 1, Len(path)) '�t�@�C��������g���q�擾
         Case "xls", "xlsx" 'Excel97-2003,Excel2007�ȍ~
            Set Office = Excel.Application
            With Office.Workbooks.Open(path.Value)
               .ExportAsFixedFormat Type:=xlTypePDF, _
               Filename:=PDF_path, Openafterpublish:=False
               .Close
            End With
            
         Case "doc", "docx" 'Word97-2003,Word2007�ȍ~
            Set Office = CreateObject("Word.Application")
            With Office.Documents.Open(path.Value)
               .ExportAsFixedFormat OutputFileName:=PDF_path, _
               ExportFormat:=17
               .Close
            End With
            Office.Quit
            
         Case "ppt", "pptx" 'Powerpoint97-2003,Powerpoint2007�ȍ~
            Set Office = CreateObject("Powerpoint.Application")
            With Office.Presentations.Open(path.Value)
               .SaveAs Filename:=PDF_path, FileFormat:=32
               .Close
            End With
            Office.Quit
      End Select
      i = i + 1
   Next
   PDFCONV = "�������܂����I"
End Function


