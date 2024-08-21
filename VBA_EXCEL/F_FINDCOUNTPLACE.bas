Attribute VB_Name = "F_FINDCOUNTPLACE"
Option Explicit

Function FINDCOUNTPLACE(‘ÎÛ As Range, _
                        ŒŸõ•¶š As Variant, _
                        Optional ‰ñ” As Long = 1)
   Dim i As Long, j As Long
   Dim myA(), A, C
   ReDim myA(‘ÎÛ.Count - 1)
   For Each C In ‘ÎÛ
      A = Split(C, ŒŸõ•¶š)
      If UBound(A) = 0 Then Exit Function
      For j = 0 To ‰ñ” - 1
         If j = ‰ñ” - 1 Then
            myA(i) = myA(i) + Len(A(j)) + 1
         Else
            myA(i) = myA(i) + Len(A(j)) + Len(ŒŸõ•¶š)
         End If
      Next j
      i = i + 1
   Next C
   FINDCOUNTPLACE = WorksheetFunction.Transpose(myA)
End Function
