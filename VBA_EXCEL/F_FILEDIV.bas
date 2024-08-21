Attribute VB_Name = "F_FILEDIV"
Option Explicit

Function FILEDIV(ëŒè€ As Range) As Variant
   Dim o As Range, i As Long, myA()
   ReDim myA(2, ëŒè€.Count - 1)
   For Each o In ëŒè€
      myA(0, i) = Left(o, InStrRev(o, "\"))
      myA(1, i) = Mid(o, InStrRev(o, "\") + 1, InStrRev(o, ".") - InStrRev(o, "\") - 1)
      myA(2, i) = Right(o, Len(o) - InStrRev(o, ".") + 1)
      i = i + 1
   Next o
   FILEDIV = WorksheetFunction.Transpose(myA)
End Function
