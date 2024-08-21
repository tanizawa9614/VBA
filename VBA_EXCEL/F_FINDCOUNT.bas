Attribute VB_Name = "F_FINDCOUNT"
Option Explicit

Function FINDCOUNT(‘ÎÛ As Range, ŒŸõ•¶š As Variant)
   Dim i As Long
   Dim myA(), C
   ReDim myA(‘ÎÛ.Count - 1)
   For Each C In ‘ÎÛ
      myA(i) = UBound(Split(C, ŒŸõ•¶š))
      i = i + 1
   Next C
   FINDCOUNT = WorksheetFunction.Transpose(myA)
End Function
