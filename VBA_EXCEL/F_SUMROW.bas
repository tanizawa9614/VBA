Attribute VB_Name = "F_SUMROW"
Option Explicit

Function SUMROW(ParamArray Target())
   Application.Volatile
   Dim A(), i As Long, j As Long
   Dim k As Long
   Dim buf
   ReDim A(UBound(Target(0).Value2) - 1)
   For k = 0 To UBound(Target)
      buf = Target(k).Value2
      For i = LBound(buf, 1) To UBound(buf, 1)
         For j = LBound(buf, 2) To UBound(buf, 2)
            A(i - 1) = A(i - 1) + buf(i, j)
         Next
      Next
   Next
   SUMROW = WorksheetFunction.Transpose(A)
End Function

