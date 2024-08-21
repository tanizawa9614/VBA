Attribute VB_Name = "F_MOVINGAVERAGE"
Option Explicit
Dim A, B, n, i As Long, j As Long
Dim buf As Double
Function MOVINGAVERAGE(”ÍˆÍ, ‹æŠÔ)
   Dim A, B, n, i As Long, j As Long
   Dim buf As Double
   Application.Volatile
   A = ”ÍˆÍ
   ReDim B(UBound(A, 1) - 1, 0)
   n = ‹æŠÔ
   If ‹æŠÔ Mod 2 = 1 Then n = n - 1
   For i = LBound(A, 1) To UBound(A, 1)
      If i >= n / 2 + 1 And i <= UBound(A, 1) - n / 2 Then
         buf = 0
         For j = 1 To n + 1
            If (j = 1 Or j = n + 1) And ‹æŠÔ Mod 2 = 0 Then
               buf = buf + 0.5 * A(i - n / 2 + j - 1, 1)
            Else
               buf = buf + A(i - n / 2 + j - 1, 1)
            End If
         Next j
         B(i - 1, 0) = buf / ‹æŠÔ
      Else
         B(i - 1, 0) = "-"
      End If
   Next i
   MOVINGAVERAGE = B
End Function
