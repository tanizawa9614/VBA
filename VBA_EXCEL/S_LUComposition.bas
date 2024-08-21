Attribute VB_Name = "S_LUComposition"
Option Explicit
Option Base 1
Sub LUComposition()
'***LU分解の必要十分条件として行列式がゼロではないこと
   Dim A, L, U
   Dim i As Long, j As Long
   Dim k As Long, n As Long
   Dim buf As Double
   Range("F:AB").Clear
   A = Range("B2").CurrentRegion
   n = UBound(A)
   ReDim L(n, n), U(n, n)
   For i = 1 To n
      For j = 1 To n
         L(i, j) = 0
         U(i, j) = 0
      Next
   Next
   For k = 1 To n
      U(1, k) = A(1, k)
      L(k, k) = 1
   Next
   For j = 1 To n
      For i = 1 To j
         U(1, j) = A(1, j)
         buf = 0
         For k = 1 To i - 1
            buf = buf + L(i, k) * U(k, j)
         Next
         U(i, j) = A(i, j) - buf
      Next
      For i = j + 1 To n
         buf = 0
         For k = 1 To j - 1
            buf = buf + L(i, k) * U(k, j)
         Next
'         If U(j, j) = 0 Then PibotSelect(A,L,
         L(i, j) = (A(i, j) - buf) / U(j, j)
      Next
   Next
   With Cells(Rows.Count, "F").End(xlUp).Offset(1).Resize(n, n)
      .Value = L
      .Offset(n + 1) = U
   End With
End Sub
