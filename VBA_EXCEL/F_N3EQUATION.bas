Attribute VB_Name = "F_N3EQUATION"
Option Explicit
Option Base 1
Function N3EQUATION(a, b, c, d)
   Dim n As Long
   Dim p, q, o, x(3), i As Long
   Dim r, buf1, buf2, buf3
   
   p = (-b ^ 2 + 3 * a * c) / 9 / a ^ 2
   q = (2 * b ^ 3 - 9 * a * b * c + 27 * a ^ 2 * d) / 54 / a ^ 3
   r = -b / 3 / a
   With WorksheetFunction
      o = .Complex(-0.5, 0.5 * Sqr(3))
      buf1 = .ImPower(q ^ 2 + p ^ 3, 0.5)
      buf2 = .ImPower(.ImSum(-q, buf1), 1 / 3)
      buf3 = .ImPower(.ImSub(-q, buf1), 1 / 3)
      x(1) = .ImSum(r, buf2, buf3)
      x(2) = .ImSum(r, .ImProduct(o, buf2), .ImProduct(.ImPower(o, 2), buf3))
      x(3) = .ImSum(r, .ImProduct(.ImPower(o, 2), buf2), .ImProduct(o, buf3))
   End With
   For i = 1 To 3
      x(i) = IMROUND(x(i), 10)
   Next
   N3EQUATION = WorksheetFunction.Transpose(x)
End Function
Function IMROUND(buf, n As Long)
   Dim nReal As Double, nImaginary As Double
   With WorksheetFunction
      nReal = Round(.ImReal(buf), n)
      nImaginary = Round(.Imaginary(buf), n)
      If nImaginary = 0 Then
         IMROUND = nReal
      Else
         IMROUND = .Complex(nReal, nImaginary)
      End If
   End With
End Function

Function IMSUM2(ParamArray myA())
   Dim buf
   For i = LBound(myA) To UBound(myA)
      If i <> UBound(myA) Then _
         buf = WorksheetFunction.ImSum(myA(i), myA(i + 1))
   Next
   IMSUM2 = buf
End Function
