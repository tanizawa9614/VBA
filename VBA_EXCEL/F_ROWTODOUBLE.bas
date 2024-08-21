Attribute VB_Name = "F_ROWTODOUBLE"
Option Explicit
Option Base 1
Function ROWTODOUBLE(”ÍˆÍ)
   Dim A, myA()
   Dim i As Long, j As Long
   On Error Resume Next
   Application.Volatile
   A = ”ÍˆÍ
   ReDim myA(UBound(A, 1) / 2, 2)
   For i = 1 To UBound(A, 1) / 2
      j = j + 1
      myA(i, 1) = A(j, 1)
      j = j + 1
      myA(i, 2) = A(j, 1)
   Next
   ROWTODOUBLE = myA
End Function
