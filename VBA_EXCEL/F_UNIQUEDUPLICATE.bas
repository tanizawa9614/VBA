Attribute VB_Name = "F_UNIQUEDUPLICATE"
Option Explicit
Function UNIQUEDUPLICATE(�͈� As Range)
   Dim A, B, i As Long
   ReDim B(�͈�.Count)
   With WorksheetFunction
      For Each A In .Unique(�͈�)
         If Not IsEmpty(A) Then
            If .CountIf(�͈�, A) > 1 Then
               B(i) = A
               i = i + 1
            End If
         End If
      Next A
      ReDim Preserve B(i - 1)
      UNIQUEDUPLICATE = .Transpose(B)
   End With
End Function
