Attribute VB_Name = "Module1"
Option Explicit

Function SPLIT3(������ As Range, _
                  ��؂蕶�� As String, _
                  Optional �e�����ŕ��� As Boolean = True, _
                  Optional ��͍폜 As Boolean = True)

   Dim A, B
   Dim C As Range, n As Long, i As Long, D As Long
   Dim C1(), l As Long, myA()
   D = ������.Count - 1
   ReDim C1(D)
   Dim mymax As Long

   For Each C In ������
      C1(l) = C.Value
      If �e�����ŕ��� = True Then
         For i = 1 To Len(��؂蕶��)
            C1(l) = Replace(C1(l), Mid(��؂蕶��, i, 1), vbTab)
         Next i
         A = Split(C1(l), vbTab)
      Else
         A = Split(C1(l), ��؂蕶��)
      End If
      
      If ��͍폜 = True Then
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

   SPLIT3 = myA
End Function
Function SPLIT2(������ As String, _
                  ��؂蕶�� As String, _
                  Optional �e�����ŕ��� As Boolean = True, _
                  Optional ��͍폜 As Boolean = True)

   Dim A, B
   Dim n As Long, i As Long
   
   If �e�����ŕ��� = True Then
      For i = 1 To Len(��؂蕶��)
         ������ = Replace(������, Mid(��؂蕶��, i, 1), vbTab)
      Next i
      ��؂蕶�� = vbTab
   End If
   
   A = Split(������, ��؂蕶��)
   
   If ��͍폜 = True Then
      B = A
      For i = 0 To UBound(A)
         If A(i) <> "" Then
            B(n) = A(i)
            n = n + 1
         End If
      Next i
         ReDim Preserve B(n - 1)
         A = B
   End If
      
   SPLIT2 = A
End Function


