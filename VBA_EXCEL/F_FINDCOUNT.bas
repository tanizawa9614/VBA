Attribute VB_Name = "F_FINDCOUNT"
Option Explicit

Function FINDCOUNT(�Ώ� As Range, �������� As Variant)
   Dim i As Long
   Dim myA(), C
   ReDim myA(�Ώ�.Count - 1)
   For Each C In �Ώ�
      myA(i) = UBound(Split(C, ��������))
      i = i + 1
   Next C
   FINDCOUNT = WorksheetFunction.Transpose(myA)
End Function
