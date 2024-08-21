Attribute VB_Name = "F_PDFCONV"
Option Explicit
Function PDFCONV(対象 As Range) As String
   Dim PDF_path As String
   Dim path As Range, i As Long
   Dim Office As Object
   Dim flag As String
   flag = MsgBox("PDFへの変換を開始しますか？", vbQuestion + vbYesNo)
   If flag = vbNo Then
      PDFCONV = "処理を中止しました"
      Exit Function
   End If
   For Each path In 対象
      PDF_path = Left(path, InStrRev(path, ".")) & "pdf"
      Select Case Mid(path, InStrRev(path, ".") + 1, Len(path)) 'ファイル名から拡張子取得
         Case "xls", "xlsx" 'Excel97-2003,Excel2007以降
            Set Office = Excel.Application
            With Office.Workbooks.Open(path.Value)
               .ExportAsFixedFormat Type:=xlTypePDF, _
               Filename:=PDF_path, Openafterpublish:=False
               .Close
            End With
            
         Case "doc", "docx" 'Word97-2003,Word2007以降
            Set Office = CreateObject("Word.Application")
            With Office.Documents.Open(path.Value)
               .ExportAsFixedFormat OutputFileName:=PDF_path, _
               ExportFormat:=17
               .Close
            End With
            Office.Quit
            
         Case "ppt", "pptx" 'Powerpoint97-2003,Powerpoint2007以降
            Set Office = CreateObject("Powerpoint.Application")
            With Office.Presentations.Open(path.Value)
               .SaveAs Filename:=PDF_path, FileFormat:=32
               .Close
            End With
            Office.Quit
      End Select
      i = i + 1
   Next
   PDFCONV = "完了しました！"
End Function


