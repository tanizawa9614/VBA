Attribute VB_Name = "Module2"
Option Explicit

Sub PDFÇ…ïœä∑()
   Dim Path  As Range, filePath As String
   Dim Word As Object
   Set Word = CreateObject("Word.Application")
   Dim i As Long
   Do While Cells(i + 1, 1) <> ""
      Set Path = Cells(i + 1, 1)
      filePath = FILEDIV(Path)(1)
      With Word.Documents.Open(Path.Value)
         .Saved = True
         Call .ChangeFileAccess(xlReadWrite)
         .ExportAsFixedFormat OutputFileName:=filePath, _
         ExportFormat:=17
         .Close
      End With
      Word.Quit
      i = i + 1
   Loop
End Sub
Function FILEDIV(ëŒè€ As Range) As Variant
   Dim A, B As String
   Dim o As Range
   Dim i As Long, j As Long
   Dim myA()
   ReDim myA(2, ëŒè€.Count - 1)
   Const C = "\"
   Const C2 = "."
   
   For Each o In ëŒè€
      A = Split(o, C)
      B = ""
      For i = 0 To UBound(A) - 1
         B = B & A(i) & C
      Next i
      myA(0, j) = B
      B = ""
      A = Split(A(UBound(A)), C2)
      For i = 0 To UBound(A) - 1
         If i <> UBound(A) - 1 Then
            B = B & A(i) & C2
         Else
            B = B & A(i)
         End If
      Next i
      myA(1, j) = B
      myA(2, j) = C2 & A(UBound(A))
      j = j + 1
   Next o
   FILEDIV = WorksheetFunction.Transpose(myA)
End Function

