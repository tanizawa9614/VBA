Attribute VB_Name = "F_MMULTEXPO"
Option Explicit

Function MMULTEXPO(�Ώۍs��, �搔 As Long)
   Dim A, i As Long, n As Long
   A = �Ώۍs��
   n = �搔
   If n = 1 Then
      MMULTEXPO = A
      Exit Function
   End If
   For i = 2 To n
      A = WorksheetFunction.MMULT(A, �Ώۍs��)
   Next
   MMULTEXPO = A
End Function
