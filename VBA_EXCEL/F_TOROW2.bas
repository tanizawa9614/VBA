Attribute VB_Name = "F_TOROW2"
Option Explicit

Function TOROW2(”ÍˆÍ, _
Optional ‹ó”’ƒZƒ‹‚ð–³Ž‹‚·‚é As Boolean = True)
   Dim colsum As Long, rowsum As Long
   Dim A, i As Long, j As Long, B()
   Dim cnt As Long
   Application.Volatile
   A = ”ÍˆÍ
   rowsum = UBound(A, 1)
   colsum = UBound(A, 2)
   ReDim B(colsum * rowsum - 1)
   For j = 1 To UBound(A, 2)
      For i = 1 To UBound(A, 1)
         If A(i, j) = "" Then
            If ‹ó”’ƒZƒ‹‚ð–³Ž‹‚·‚é = False Then
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
   TOROW2 = WorksheetFunction.Transpose(B)
End Function
