Attribute VB_Name = "F_LU"
Option Explicit
Option Base 1
Function LU(A0, Optional 表示 As Long = 0)
   Dim L, U, n As Long, A
   Dim i As Long, j As Long, k As Long
   Dim tmp As Double
   
   A = A0
   n = UBound(A)
   L = LU_ZERO(n)
   U = LU_MUNIT(n)
   For k = 1 To n
      For i = k To n
         tmp = 0
         For j = 1 To k - 1
            tmp = tmp + L(i, j) * U(j, k)
         Next
         L(i, k) = A(i, k) - tmp
      Next
      For j = k + 1 To n
         tmp = 0
         For i = 1 To k - 1
            tmp = tmp + L(k, i) * U(i, j)
         Next
         If L(k, k) = 0 Then L(k, k) = 10 ^ -10
         U(k, j) = (A(k, j) - tmp) / L(k, k)
      Next
   Next
   
   If 表示 = 0 Then
      LU = LU_MJOIN(L, U)
   ElseIf 表示 = 1 Then
      LU = L
   ElseIf 表示 = 2 Then
      LU = U
   End If
End Function
Private Function LU_MJOIN(L, U)
   Dim A, nC As Long
   Dim i As Long, j As Long
   A = L
   nC = UBound(U, 2)
   ReDim Preserve A(UBound(A, 1), UBound(A, 2) + nC)
   For i = 1 To UBound(A, 1)
      For j = UBound(L, 2) + 1 To nC + UBound(L, 2)
         A(i, j) = U(i, j - UBound(L, 2))
      Next
   Next
   LU_MJOIN = A
End Function
Private Function LU_MUNIT(n As Long)
   Dim A(), i As Long, j As Long
   ReDim A(n, n)
   For i = 1 To n
      For j = 1 To n
         If i = j Then
            A(i, j) = 1
         Else
            A(i, j) = 0
         End If
      Next
   Next
   LU_MUNIT = A
End Function
Private Function LU_ZERO(n As Long)
   Dim A(), i As Long, j As Long
   ReDim A(n, n)
   For i = 1 To n
      For j = 1 To n
         A(i, j) = 0
      Next
   Next
   LU_ZERO = A
End Function
