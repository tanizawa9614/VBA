Attribute VB_Name = "F_FINDCOUNTPLACE"
Option Explicit

Function FINDCOUNTPLACE(�Ώ� As Range, _
                        �������� As Variant, _
                        Optional �� As Long = 1)
   Dim i As Long, j As Long
   Dim myA(), A, C
   ReDim myA(�Ώ�.Count - 1)
   For Each C In �Ώ�
      A = Split(C, ��������)
      If UBound(A) = 0 Then Exit Function
      For j = 0 To �� - 1
         If j = �� - 1 Then
            myA(i) = myA(i) + Len(A(j)) + 1
         Else
            myA(i) = myA(i) + Len(A(j)) + Len(��������)
         End If
      Next j
      i = i + 1
   Next C
   FINDCOUNTPLACE = WorksheetFunction.Transpose(myA)
End Function
