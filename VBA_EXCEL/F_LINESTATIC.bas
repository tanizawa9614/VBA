Attribute VB_Name = "F_LINESTATIC"
Option Explicit

Function LINESTATIC(���m��y As Range, ���m��x As Range)
   Dim A(9, 1), y, x
   y = ���m��y.Value
   x = ���m��x.Value
   A(0, 0) = "��A����"
   A(1, 0) = "a"
   A(2, 0) = "b"
   A(3, 0) = "���֌W�� R"
   A(4, 0) = "����W�� R^2"
   A(5, 0) = "x�̕���"
   A(6, 0) = "y�̕���"
   A(7, 0) = "x�̕΍����a"
   A(8, 0) = "y�̕΍����a"
   A(9, 0) = "x,y�̕΍��Ϙa"
   A(0, 1) = "y=ax+b"
   
   A(1, 1) = WorksheetFunction.Slope(y, x)
   A(2, 1) = WorksheetFunction.Intercept(y, x)
   A(3, 1) = WorksheetFunction.Correl(y, x)
   A(4, 1) = A(3, 1) ^ 2
   A(5, 1) = WorksheetFunction.Average(x)
   A(6, 1) = WorksheetFunction.Average(y)
   A(7, 1) = WorksheetFunction.SumSq(x) - UBound(x, 1) * A(5, 1) ^ 2
   A(8, 1) = WorksheetFunction.SumSq(y) - UBound(y, 1) * A(6, 1) ^ 2
   
   

End Function
