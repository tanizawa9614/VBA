Attribute VB_Name = "F_UNIQUEDUP"
Option Explicit

Function UNIQUEDUP(”ÍˆÍ As Range)
   Dim A, B(), i As Long
   ReDim B(”ÍˆÍ.Count)
   With WorksheetFunction
      For Each A In .Unique(”ÍˆÍ)
         If IsEmpty(A) Then GoTo L1
         If .CountIf(”ÍˆÍ, A) > 1 Then
            B(i) = A
            i = i + 1
         End If
L1:
      Next A
      ReDim Preserve B(i - 1)
      UNIQUEDUP = .Transpose(B)
   End With
End Function

