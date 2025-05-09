Attribute VB_Name = "F_FINDCOUNTPLACE"
Option Explicit

Function FINDCOUNTPLACE(対象 As Range, _
                        検索文字 As Variant, _
                        Optional 回数 As Long = 1)
   Dim i As Long, j As Long
   Dim myA(), A, C
   ReDim myA(対象.Count - 1)
   For Each C In 対象
      A = Split(C, 検索文字)
      If UBound(A) = 0 Then Exit Function
      For j = 0 To 回数 - 1
         If j = 回数 - 1 Then
            myA(i) = myA(i) + Len(A(j)) + 1
         Else
            myA(i) = myA(i) + Len(A(j)) + Len(検索文字)
         End If
      Next j
      i = i + 1
   Next C
   FINDCOUNTPLACE = WorksheetFunction.Transpose(myA)
End Function
