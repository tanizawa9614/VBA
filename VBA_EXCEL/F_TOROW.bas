Attribute VB_Name = "F_TOROW"
Option Explicit

Function TOROW(�͈�, _
Optional �󔒃Z���𖳎����� As Boolean = True)
   Dim colsum As Long, rowsum As Long
   Dim A, i As Long, j As Long, B()
   Dim cnt As Long
   Application.Volatile
   A = �͈�
   rowsum = UBound(A, 1)
   colsum = UBound(A, 2)
   ReDim B(colsum * rowsum - 1)
   For j = 1 To UBound(A, 2)
      For i = 1 To UBound(A, 1)
         If A(i, j) = "" Then
            If �󔒃Z���𖳎����� = False Then
               B(cnt) = A(i, j)
               cnt = cnt + 1
            End If
         Else
            B(cnt) = A(i, j)
            cnt = cnt + 1
         End If
      Next
   Next
   ReDim Preserve B(cnt - 1)
   TOROW = WorksheetFunction.Transpose(B)
End Function
