Attribute VB_Name = "F_SPLIT2"
Option Explicit
Function SPLIT2(•¶š—ñ, _
               ‹æØ‚è•¶š, _
               Optional ‹ó‚Ííœ As Boolean = True, _
               Optional ‘ã‘Ö•¶š As String = "")

   Dim A, B, i As Long, j As Long
   Dim C, n As Long, D As Long
   Dim l As Long, myA(), Delimiter, buf
   Dim mymax As Long
   
   Application.Volatile
   If IsArray(•¶š—ñ) Then
      C = •¶š—ñ
      D = UBound(C, 1) - 1
   Else
      D = 0
   End If
   ReDim C1(D)
   
   For Each C In •¶š—ñ
      Delimiter = ‹æØ‚è•¶š
      If IsArray(Delimiter) Then
         For Each buf In Delimiter
            C = Replace(C, buf, vbTab)
         Next
      Else
         C = Replace(C, Delimiter, vbTab)
      End If
      A = Split(C, vbTab)
      
      If ‹ó‚Ííœ Then
         B = A
         n = 0
         For i = 0 To UBound(A)
            If A(i) <> "" Then
               B(n) = A(i)
               n = n + 1
            End If
         Next i
         ReDim Preserve B(n - 1)
         A = B
      End If
      If mymax <> WorksheetFunction.Max(mymax, UBound(A)) Then
         mymax = WorksheetFunction.Max(mymax, UBound(A))
         ReDim Preserve myA(D, mymax)
      End If
      
      For i = 0 To UBound(A)
         myA(l, i) = A(i)
      Next i
      l = l + 1
   Next C
   
   For i = 0 To UBound(myA, 1)
      For j = 0 To UBound(myA, 2)
         If IsEmpty(myA(i, j)) Then myA(i, j) = ‘ã‘Ö•¶š
      Next
   Next

   SPLIT2 = myA
End Function


