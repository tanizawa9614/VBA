Attribute VB_Name = "F_UNIQUEMULTI"
Option Explicit
Option Base 1
Function UNIQUEMULTI(”z—ñ As Variant)
   Dim A, B, C(), D, i As Long, n As Long
   Dim t As Double, buf As String
   t = Timer
   With WorksheetFunction
      A = .Transpose(.Unique(”z—ñ))
      B = .Transpose(.Unique(”z—ñ, , True))
   End With
   
   n = UBound(A)
   ReDim Preserve A(n + UBound(B))
   For i = 1 To UBound(B)
      A(i + n) = B(i)
   Next i
   With WorksheetFunction
      UNIQUEMULTI = .Transpose(.Unique(A, True, True))
   End With
   
End Function
