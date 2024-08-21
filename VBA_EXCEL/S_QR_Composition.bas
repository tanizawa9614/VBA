Attribute VB_Name = "S_QR_Composition"
Option Explicit
Option Base 1
Sub QR_Composition()
   Dim i As Long, j As Long, k As Long, cnt As Long
   Dim Q(), R(), x(), y(), u(), H(), R0
   Dim A, A0, n As Long, lmd(), P()
   
   A0 = Range("B2").CurrentRegion
   A = A0
   n = UBound(A)
   ReDim B(n, 1), lmd(1, n), P(n, n)
   ReDim Q(n, n), R(n, n), H(n, n)
   
   
   For cnt = 1 To 2000
      R = A
      Q = MUNIT(n)
      For k = 1 To n - 1
         ReDim x(n, 1), y(n, 1), u(n, 1)
         x = M_Initialize(x)
         y = M_Initialize(y)
         For j = k To n
            x(j, 1) = R(j, k)
         Next
         y(k, 1) = MABS(x)
         u = MCONST_PRODUCT(MSUB(x, y), MABS(MSUB(x, y)) ^ -1)
         With WorksheetFunction
            H = MSUB(MUNIT(n), MCONST_PRODUCT(.MMult(u, .Transpose(u)), 2))
            R = .MMult(H, R)
            Q = .MMult(Q, .Transpose(H))
         End With
      Next k
      A = WorksheetFunction.MMult(R, Q)
      If cnt >= 2 Then
         If Žû‘©”»’è(R, R0) Then Exit For
      End If
      R0 = R
   Next cnt
   
   For i = 1 To n
      lmd(1, i) = R(i, i)
   Next
   
   P = reverse_iteration(A0, lmd)
   
   With Range("F2").Resize(n, n)
      .Value = Q
      .Offset(4) = R
   End With
End Sub
Function reverse_iteration(A, lmd)
   Dim i As Long, j As Long, k As Long
   Dim n As Long, B, P
   Dim y()
   n = UBound(A, 1)
   ReDim y(n, 1)
   '‰Šú’lÝ’è
   For i = 1 To n
      y(i, 1) = 1
   Next
   Dim buf
   For k = 1 To n
      With WorksheetFunction
         B = .MInverse(MSUB(A, MCONST_PRODUCT(MUNIT(n), lmd(1, k))))
         For i = 1 To 100
            y = .MMult(B, y)
            buf = MFINDMIN(y)
            For j = 1 To n
               y(j, 1) = y(j, 1) / buf
            Next
         Next
      End With
      If k = 1 Then
         P = y
      Else
         P = MJOIN(P, y)
      End If
   Next
   reverse_iteration = P
End Function
Function MFINDMIN(n) As Double
   Dim i As Long, n_max As Double
   Dim n_min As Double
   With WorksheetFunction
      n_max = .Max(n)
      n_min = .Min(n)
      If n_max >= Abs(n_min) Then
         MFINDMIN = n_max
      Else
         MFINDMIN = n_min
      End If
   End With
End Function
Function MJOIN(Ar1, Ar2)
   Dim i As Long, j As Long, cnt As Long
   Dim nc1 As Long, nc2 As Long, nr As Long
   Dim Ar()
   nr = UBound(Ar1, 1)
   nc1 = UBound(Ar1, 2)
   nc2 = UBound(Ar2, 2)
   ReDim Ar(nr, nc1 + nc2)
   For i = 1 To nr
      For j = 1 To nc1
         Ar(i, j) = Ar1(i, j)
      Next
   Next
   For i = 1 To nr
      For j = 1 To nc2
         Ar(i, j + nc1) = Ar2(i, j)
      Next
   Next
   MJOIN = Ar
End Function
   
Function Žû‘©”»’è(l, M) As Boolean
   Dim i As Long, j As Long, cnt As Long
   Dim buf As Double
   For i = 1 To UBound(l, 1)
      If Abs(l(i, i) - M(i, i)) / Abs(l(i, i)) <= 10 ^ -7 Then _
            cnt = cnt + 1
   Next
   If cnt = UBound(l, 1) Then
      Žû‘©”»’è = True
   Else
      Žû‘©”»’è = False
   End If
End Function
Function M_Initialize(A1)
   Dim i As Long, j As Long, A, C()
   A = A1
   ReDim C(UBound(A, 1), UBound(A, 2))
   For i = LBound(A, 1) To UBound(A, 1)
      For j = LBound(A, 2) To UBound(A, 2)
         C(i, j) = 0
      Next
   Next
   M_Initialize = C
End Function

Function MCONST_PRODUCT(A1, D)
   Dim i As Long, j As Long, C(), A
   A = A1
   ReDim C(UBound(A, 1), UBound(A, 2))
   For i = LBound(A, 1) To UBound(A, 1)
      For j = LBound(A, 2) To UBound(A, 2)
         C(i, j) = A(i, j) * D
      Next
   Next
   MCONST_PRODUCT = C
End Function

Function MSUM(A1, B1)
   Dim i As Long, j As Long, A, B, C()
   A = A1
   B = B1
   ReDim C(UBound(A, 1), UBound(A, 2))
   For i = LBound(A, 1) To UBound(A, 1)
      For j = LBound(A, 2) To UBound(A, 2)
         C(i, j) = A(i, j) + B(i, j)
      Next
   Next
   MSUM = C
End Function

Function MSUB(A1, B1)
   Dim i As Long, j As Long, A, B, C()
   A = A1
   B = B1
   ReDim C(UBound(A, 1), UBound(A, 2))
   For i = LBound(A, 1) To UBound(A, 1)
      For j = LBound(A, 2) To UBound(A, 2)
         C(i, j) = A(i, j) - B(i, j)
      Next
   Next
   MSUB = C
End Function
Function MABS(vector1)
   Dim i As Long, buf As Double, vector
   vector = vector1
   For i = LBound(vector, 1) To UBound(vector, 1)
      buf = buf + vector(i, 1) ^ 2
   Next
   MABS = Sqr(buf)
End Function
Function MUNIT(n As Long)
   Dim i As Long, j As Long, A()
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
   MUNIT = A
End Function

