Attribute VB_Name = "F_�f��������2"
Option Explicit

Function PRIMEFACTO(�l As Long)
   Dim i As Long, i_max As Long
   Dim n As Long, A, j As Long
   n = �l
   i_max = Int(Sqr(n))
   ReDim A(i_max)
   For i = 2 To i_max
      Do
         If n Mod i = 0 Then
            A(j) = i
            n = n / i
            i_max = Int(Sqr(n))
            j = j + 1
         Else
            Exit Do
         End If
      Loop
   Next
   If n <> 1 Then A(j) = n
   PRIMEFACTO = Join(A)
End Function
