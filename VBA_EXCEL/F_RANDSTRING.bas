Attribute VB_Name = "F_RANDSTRING"
Option Explicit

Function RANDSTRING(�s As Long, _
   Optional �� As Long = 1, _
   Optional ������ As Long = 1)
   
   Dim Unicode_�� As Long
   Dim Unicode_�� As Long
   Dim A(), i As Long, j As Long
   Dim k As Long
   ReDim A(�s - 1, �� - 1)
   
   Application.Volatile
   Unicode_�� = 12354
   Unicode_�� = 12435
   For i = 0 To �s - 1
      For j = 0 To �� - 1
         For k = 1 To ������
            With WorksheetFunction
               A(i, j) = A(i, j) & .Unichar(.RandBetween(Unicode_��, Unicode_��))
            End With
         Next
      Next
   Next
   RANDSTRING = A
End Function


