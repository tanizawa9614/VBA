Attribute VB_Name = "F_MMULTEXPO"
Option Explicit

Function MMULTEXPO(‘ÎÛs—ñ, æ” As Long)
   Dim A, i As Long, n As Long
   A = ‘ÎÛs—ñ
   n = æ”
   If n = 1 Then
      MMULTEXPO = A
      Exit Function
   End If
   For i = 2 To n
      A = WorksheetFunction.MMULT(A, ‘ÎÛs—ñ)
   Next
   MMULTEXPO = A
End Function
