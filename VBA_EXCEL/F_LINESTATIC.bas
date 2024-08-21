Attribute VB_Name = "F_LINESTATIC"
Option Explicit

Function LINESTATIC(Šù’m‚Ìy As Range, Šù’m‚Ìx As Range)
   Dim A(9, 1), y, x
   y = Šù’m‚Ìy.Value
   x = Šù’m‚Ìx.Value
   A(0, 0) = "‰ñ‹A’¼ü"
   A(1, 0) = "a"
   A(2, 0) = "b"
   A(3, 0) = "‘ŠŠÖŒW” R"
   A(4, 0) = "Œˆ’èŒW” R^2"
   A(5, 0) = "x‚Ì•½‹Ï"
   A(6, 0) = "y‚Ì•½‹Ï"
   A(7, 0) = "x‚Ì•Î·“ñæ˜a"
   A(8, 0) = "y‚Ì•Î·“ñæ˜a"
   A(9, 0) = "x,y‚Ì•Î·Ï˜a"
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
