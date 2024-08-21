Attribute VB_Name = "F_EIGENVALUE3"
Option Explicit
Option Base 1
Function EIGENVALUE(çsóÒ)
   Dim i As Long, j As Long, cnt As Long
   Dim A, B(), n As Long, lmd(), cnt2 As Long, P()
   Dim x_max As Double, x_min As Double, x_abs As Double
   Dim tmp_lmd As Double
   A = çsóÒ
   n = UBound(A)
   ReDim B(n, 1), lmd(1, n), P(n, n)
   
   For cnt2 = 1 To n
      For i = 1 To n
         A(i, i) = A(i, i) - x_abs
      Next
      For i = 1 To n
         B(i, 1) = 1
      Next
      For cnt = 1 To 100
         B = WorksheetFunction.MMult(A, B)
         x_max = 0
         x_min = 0
         For i = 1 To n
            x_max = WorksheetFunction.Max(B(i, 1), x_max)
            x_min = WorksheetFunction.Min(B(i, 1), x_min)
         Next
         If x_max >= Abs(x_min) Then
            x_abs = x_max
         Else
            x_abs = x_min
         End If
         For i = 1 To n
            B(i, 1) = B(i, 1) / x_abs
         Next
      Next
      lmd(1, cnt2) = x_abs + WorksheetFunction.Sum(lmd)
      For i = 1 To n
         P(i, cnt2) = B(i, 1)
      Next
   Next cnt2
   EIGENVALUE = JOINARRAY(lmd, P, n)
End Function
Private Function JOINARRAY(Ar1, Ar2, n)
   Dim Ar(), i As Long, j As Long
   ReDim Ar(3 + n, n)
   For i = 1 To n
      Ar(1, i) = "É…" & i
      Ar(2, i) = Ar1(1, i)
      Ar(3, i) = "u" & i
   Next
   For j = 1 To n
      For i = 1 To n
         Ar(i - 1 + 4, j) = Ar2(i, j)
      Next
   Next
   JOINARRAY = Ar
End Function
