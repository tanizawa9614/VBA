Attribute VB_Name = "F_ëfàˆêîï™â"
Option Explicit

Function PRIMEFACTO(íl As Long)
   Dim i As Long, i_max As Long
   Dim n As Long, A As String, j As Long
   
   n = íl
   i_max = Int(Sqr(n))
   
   For i = 2 To i_max
      j = 0
      Do
         If n Mod i = 0 Then
            n = n / i
            j = j + 1
         Else
            Exit Do
         End If
      Loop
      If j = 1 Then
         A = A & "*" & i
      ElseIf j <> 0 Then
         A = A & "*" & i & "^" & j
      End If
      i_max = Int(Sqr(n))
   Next i
   
   If n <> 1 Then A = A & "*" & n
   A = Mid(A, 2, Len(A))
   PRIMEFACTO = íl & "=" & A
End Function
