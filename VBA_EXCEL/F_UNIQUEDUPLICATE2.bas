Attribute VB_Name = "F_UNIQUEDUPLICATE"
Option Explicit
Option Base 1
Function UNIQUEDUPLICATE(�z�� As Variant) As Variant
   Dim A, B, n As Long, i As Long
   On Error Resume Next
   With WorksheetFunction
      A = .Transpose(.Unique(�z��))
      B = .Transpose(.Unique(�z��, , True))
   End With
   n = UBound(A)
   ReDim Preserve A(UBound(A) + UBound(B))
   For i = 1 To UBound(B)
      A(i + n) = B(i)
   Next i
   With WorksheetFunction
      UNIQUEDUPLICATE = .Transpose(.Unique(A, True, True))
   End With
End Function
