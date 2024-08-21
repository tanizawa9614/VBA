Attribute VB_Name = "F_CONCAT2"
Option Explicit

Function CONCAT2(ParamArray Target())
   Application.Volatile
   Dim A(), i As Long, j As Long
   Dim k As Long
   Dim buf
   buf = 1
   For k = 0 To UBound(Target)
      If IsArray(Target(k).Value2) Then
         If buf < UBound(Target(k).Value2) Then
            buf = UBound(Target(k).Value2)
         End If
      End If
   Next
   ReDim A(buf - 1)
   For k = 0 To UBound(Target)
      buf = Target(k)
      If IsArray(buf) Then
         For i = LBound(buf, 1) To UBound(buf, 1)
            For j = LBound(buf, 2) To UBound(buf, 2)
               A(i - 1) = A(i - 1) & buf(i, j)
            Next
         Next
      Else
         For i = LBound(A) To UBound(A)
            A(i) = A(i) & buf
         Next
      End If
   Next
   CONCAT2 = WorksheetFunction.Transpose(A)
End Function
