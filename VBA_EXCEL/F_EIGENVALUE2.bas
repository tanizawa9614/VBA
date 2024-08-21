Attribute VB_Name = "F_EIGENVALUE2"
Option Explicit
Option Base 1
Dim n As Long, A, l, U, P, lmd
Function EIGENVALUE2(Ar)
   A = Ar
   n = UBound(A, 1)
   ReDim l(n, n), U(n, n)
   ReDim P(n, n), lmd(n, 1)
   If n = 2 Then
      n_2 (A)
      EIGENVALUE2 = JOINARRAY(lmd, P)
   ElseIf n = 3 Then
      n_3
      EIGENVALUE2 = MAKEARRAY(A)
   End If
End Function
Private Function n_2(myA)
   Dim B As Double, C As Double
   Dim i As Long
   B = -tr(myA)
   C = myA(1, 1) * myA(2, 2) - myA(1, 2) * myA(2, 1)
   lmd(1, 1) = 0.5 * (-B - Sqr(B ^ 2 - 4 * C))
   lmd(2, 1) = 0.5 * (-B + Sqr(B ^ 2 - 4 * C))
   For i = 1 To 2
      P(i, 2) = 1
      P(i, 1) = -myA(1, 2) / (myA(1, 1) - lmd(i, 1))
   Next
End Function
Private Function n_3()
   Dim i As Long
   With WorksheetFunction
      For i = 1 To 1000
         Call LU_Decomposition
         A = .MMult(U, l)
      Next
      For i = 1 To n
         lmd(i, 1) = A(i, i)
      Next
   End With
End Function
Private Sub LU_Decomposition()
   Dim i As Long, j As Long
   ReDim l(n, n), U(n, n)
   For i = 1 To n
      For j = 1 To n
         l(i, j) = 0
         U(i, j) = 0
         If i = j Then l(i, i) = 1
      Next
   Next
   U(1, 1) = A(1, 1)
   U(1, 2) = A(1, 2)
   U(1, 3) = A(1, 3)
   l(2, 1) = A(2, 1) / U(1, 1)
   l(3, 1) = A(3, 1) / U(1, 1)
   U(2, 2) = A(2, 2) - l(2, 1) * U(1, 2)
   U(2, 3) = A(2, 3) - l(2, 1) * U(1, 3)
   l(3, 2) = (A(3, 2) - l(3, 1) * U(1, 2)) / U(2, 2)
   U(3, 3) = A(3, 3) - l(3, 1) * U(1, 3) - l(3, 2) * U(2, 3)
End Sub
Private Function tr(Ar1)
   Dim buf As Double
   Dim i As Long
   For i = 1 To n
      buf = buf + Ar1(i, i)
   Next
   tr = buf
End Function
Private Function JOINARRAY(Ar1, Ar2)
   Dim Ar(), i As Long, j As Long
   ReDim Ar(5, 2)
   For i = 1 To n
      Ar(1, i) = "É…" & i
      Ar(2, i) = Ar1(i, 1)
      Ar(3, i) = "u" & i & "(vector)"
   Next
   For j = 1 To n
      For i = 1 To n
         Ar(i - 1 + 4, j) = Ar2(i, j)
      Next
   Next
   JOINARRAY = Ar
End Function
Private Function MAKEARRAY(Ar1)
   Dim Ar(), i As Long, j As Long
   ReDim Ar(2, 3)
   For i = 1 To n
      Ar(1, i) = "É…" & i
      Ar(2, i) = Ar1(i, i)
   Next
   MAKEARRAY = Ar
End Function

