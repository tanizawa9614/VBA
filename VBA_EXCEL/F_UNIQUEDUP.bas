Attribute VB_Name = "F_UNIQUEDUP"
Option Explicit

Function UNIQUEDUP(�͈� As Range)
   Dim A, B(), i As Long
   ReDim B(�͈�.Count)
   With WorksheetFunction
      For Each A In .Unique(�͈�)
         If IsEmpty(A) Then GoTo L1
         If .CountIf(�͈�, A) > 1 Then
            B(i) = A
            i = i + 1
         End If
L1:
      Next A
      ReDim Preserve B(i - 1)
      UNIQUEDUP = .Transpose(B)
   End With
End Function

