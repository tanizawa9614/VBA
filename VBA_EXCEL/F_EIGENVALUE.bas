Attribute VB_Name = "F_EIGENVALUE"
Option Explicit
Option Base 1
Function EIGENVALUE(A0)
   Dim A, lmd(), P, R, i As Long, n As Long
   A = A0
   n = UBound(A)
   ReDim lmd(1, n)
'   R = QR_Composition_GramSchmidt(A) 'QR分解_ハウスホルダー変換
   R = QR_Composition_HouseHolder(A) 'QR分解_ハウスホルダー変換
   For i = 1 To n
      lmd(1, i) = R(i, i)
'      MsgBox WorksheetFunction.MDeterm(MSUB(A0, MCONST_PRODUCT(MUNIT(n), lmd(1, i))))
   Next
   P = Reverse_Iteration(A, lmd) '逆反復法
   EIGENVALUE = JOINARRAY(lmd, P)
End Function
Private Function QR_Composition_GramSchmidt(A0)
   Dim i As Long, j As Long, k As Long, cnt As Long
   Dim u, q_comp, Q, R, R0
   Dim A, n As Long, q_sum()
   A = A0
   n = UBound(A)
   On Error GoTo myErr
   ReDim q_sum(n, 1)
   For cnt = 1 To 2000
      For i = 1 To n
         q_sum = M_Initialize(q_sum)
         For j = 1 To i - 1
            q_sum = MSUM(q_sum, MCONST_PRODUCT(q_comp, INNER_PRODUCT(MCHOOSECOL(A, i), MCHOOSECOL(Q, j))))
         Next
         u = MSUB(MCHOOSECOL(A, i), q_sum)
         q_comp = MCONST_PRODUCT(u, MABS(u) ^ -1)
         If i = 1 Then
            Q = q_comp
         Else
            Q = MJOIN(Q, q_comp)
         End If
      Next
      With WorksheetFunction
         R = .MMult(.Transpose(Q), A)
      End With
      If cnt >= 2 Then
         If 収束判定(R, R0) Then
            Exit For
         End If
      End If
      A = WorksheetFunction.MMult(Q, R)
'      A = WorksheetFunction.MMult(R, Q)
      R0 = R '収束判定用にR0を更新
   Next cnt
myErr:
   If Err.Number > 0 Then
      Stop
   End If
   QR_Composition_GramSchmidt = R
End Function

Private Function QR_Composition_HouseHolder(A0)
   Dim i As Long, j As Long, k As Long, cnt As Long
   Dim Q(), R(), x(), y(), u(), H(), R0
   Dim A, n As Long
   
   A = A0
   n = UBound(A)
   ReDim B(n, 1), lmd(1, n)
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
         If MABS(MSUB(x, y)) = 0 Then GoTo L1
         u = MCONST_PRODUCT(MSUB(x, y), MABS(MSUB(x, y)) ^ -1)
         With WorksheetFunction
            H = MSUB(MUNIT(n), MCONST_PRODUCT(.MMult(u, .Transpose(u)), 2))
            R = .MMult(H, R)
            Q = .MMult(Q, .Transpose(H))
         End With
L1:   Next k
      If cnt >= 2 Then
         If 収束判定(R, R0) Then
            Exit For
         End If
      End If
'      A = WorksheetFunction.MMult(Q, R)
      A = WorksheetFunction.MMult(R, Q)
      R0 = R '収束判定用にR0を更新
   Next cnt
   A = WorksheetFunction.MMult(R, Q)
   QR_Composition_HouseHolder = A
End Function
Private Function Reverse_Iteration(A, lmd) '逆反復法
   Dim i As Long, j As Long, k As Long
   Dim n As Long, B, P
   Dim y()
   n = UBound(A, 1)
   ReDim y(n, 1)
   '初期値設定
   For i = 1 To n
      y(i, 1) = 1
   Next
   Dim buf
   For k = 1 To n
      With WorksheetFunction
         B = MSUB(A, MCONST_PRODUCT(MUNIT(n), lmd(1, k) + 10 ^ -8))
'         MsgBox .MDeterm(B)
         B = .MInverse(B)
         For i = 1 To 1000
            y = .MMult(B, y)
            buf = MABS(y)
            For j = 1 To n
               y(j, 1) = y(j, 1) / buf
            Next
         Next
      End With
      buf = MFINDMAX(y)
      For j = 1 To n
         y(j, 1) = y(j, 1) / buf
      Next
      If k = 1 Then
         P = y
      Else
         P = MJOIN(P, y)
      End If
   Next
   Reverse_Iteration = P
End Function
Function MCHOOSECOL(A0, j)
   Dim i As Long, A, B
   A = A0
   ReDim B(UBound(A), 1)
   For i = 1 To UBound(A)
      B(i, 1) = A(i, j)
   Next
   MCHOOSECOL = B
End Function
Function INNER_PRODUCT(v0, w0)
   Dim i As Long, buf As Double, v, w
   v = v0
   w = w0
   For i = 1 To UBound(v)
      buf = buf + v(i, 1) * w(i, 1)
   Next
   INNER_PRODUCT = buf
End Function
Private Function JOINARRAY(Ar1, Ar2) '出力用配列操作
   Dim Ar(), i As Long, j As Long
   Dim n As Long
   n = UBound(Ar2)
   ReDim Ar(3 + n, n)
   For i = 1 To n
      Ar(1, i) = "λ" & i
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
Function MFINDMAX(n0) As Double
   Dim i As Long, n_max As Double, n
   Dim n_min As Double
   n = n0
   With WorksheetFunction
      n_max = .Max(n)
      n_min = .Min(n)
      If n_max <= Abs(n_min) Then
         MFINDMAX = n_min
      Else
         MFINDMAX = n_max
      End If
   End With
End Function
Function MJOIN(Ar01, Ar02)
   Dim i As Long, j As Long, cnt As Long
   Dim nc1 As Long, nc2 As Long, nr As Long
   Dim Ar(), Ar1, Ar2
   Ar1 = Ar01
   Ar2 = Ar02
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
Private Function 収束判定(l, M) As Boolean
   Dim i As Long, j As Long, cnt As Long
   Dim buf As Double
   For i = 1 To UBound(l, 1)
      If Abs(l(i, i)) = 0 Then
         cnt = cnt + 1
      ElseIf Abs(l(i, i) - M(i, i)) / Abs(l(i, i)) <= 10 ^ -10 Then
         cnt = cnt + 1
      End If
   Next
   If cnt = UBound(l, 1) Then
      収束判定 = True
   Else
      収束判定 = False
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
Private Function MUNIT(n As Long)
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

