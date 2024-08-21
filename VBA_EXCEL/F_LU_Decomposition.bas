Attribute VB_Name = "F_LU_Decomposition"
Option Explicit
Option Base 1
Dim n As Long, A, L, U, P, lmd
Function EIGENVALUE(Ar)
   A = Ar
   n = UBound(A, 1)
   ReDim L(n, n), U(n, n)
   ReDim P(n, n), lmd(n, 1)
   If n = 2 Then
      n_2 (A)
      EIGENVALUE = JOINARRAY(lmd, P)
   ElseIf n = 3 Then
      n_3
      EIGENVALUE = MAKEARRAY(A)
   End If
End Function
Private Function n_2(myA)
   Dim B As Double, C As Double
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
         A = .MMult(U, L)
      Next
      For i = 1 To n
         lmd(i, 1) = A(i, i)
      Next
   End With
End Function

Private Sub LU_Decomposition()
   Dim i As Long, j As Long
   ReDim L(n, n), U(n, n)
   For i = 1 To n
      For j = 1 To n
         L(i, j) = 0
         U(i, j) = 0
         If i = j Then L(i, i) = 1
      Next
   Next
   U(1, 1) = A(1, 1)
   U(1, 2) = A(1, 2)
   U(1, 3) = A(1, 3)
   L(2, 1) = A(2, 1) / U(1, 1)
   L(3, 1) = A(3, 1) / U(1, 1)
   U(2, 2) = A(2, 2) - L(2, 1) * U(1, 2)
   U(2, 3) = A(2, 3) - L(2, 1) * U(1, 3)
   L(3, 2) = (A(3, 2) - L(3, 1) * U(1, 2)) / U(2, 2)
   U(3, 3) = A(3, 3) - L(3, 1) * U(1, 3) - L(3, 2) * U(2, 3)
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
   ReDim Ar(n, 2 * (1 + n))
   For i = 1 To n
      Ar(i, 1) = "λ" & i & "="
      Ar(i, 2) = Ar1(i, 1)
      Ar(1, 3 + 2 * (i - 1)) = "固有ベクトル"
      Ar(2, 3 + 2 * (i - 1)) = "u" & i
   Next
   For j = 1 To n
      For i = 1 To n
         Ar(i, 4 + 2 * (j - 1)) = Ar2(i, j)
      Next
   Next
   JOINARRAY = Ar
End Function
Private Function MAKEARRAY(Ar1)
   Dim Ar(), i As Long, j As Long
   ReDim Ar(n, 2)
   For i = 1 To n
      Ar(i, 1) = "λ" & i & "="
      Ar(i, 2) = Ar1(i, 1)
   Next
   MAKEARRAY = Ar
End Function

