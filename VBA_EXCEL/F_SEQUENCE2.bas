Attribute VB_Name = "F_SEQUENCE2"
Option Explicit

Function SEQUENCE2(�J�n As Double, _
      �I�� As Double, _
      �ڐ��� As Variant, _
      Optional _
      ����� As Boolean = False)
   Dim row_count As Long
   Dim col_count As Long
      
   If �ڐ��� = 0 Then End
   If ����� Then
      row_count = 1
      col_count = (�I�� - �J�n) / �ڐ��� + 1
   Else
      row_count = (�I�� - �J�n) / �ڐ��� + 1
      col_count = 1
   End If
   SEQUENCE2 = WorksheetFunction.Sequence(row_count, col_count, �J�n, �ڐ���)
End Function
