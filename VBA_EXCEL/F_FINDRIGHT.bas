Attribute VB_Name = "F_FINDRIGHT"
Option Explicit

Function FINDRIGHT(�Ώ� As Range, �������� As Variant)
   Dim i As Long, A As Object, myA(), j As Long
   ReDim myA(�Ώ�.Count - 1)
   On Error Resume Next
   �������� = Left(��������, Len(��������))
   For Each A In �Ώ�
      For i = Len(A) To 1 Step -1
         If Mid(A, i, 1) = �������� Then
            myA(j) = i
            Exit For
         End If
      Next i
      j = j + 1
   Next A
   FINDRIGHT = WorksheetFunction.Transpose(myA)
End Function
