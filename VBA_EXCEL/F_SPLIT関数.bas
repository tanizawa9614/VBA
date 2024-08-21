Attribute VB_Name = "Module1"
Option Explicit

Function SPLIT3(文字列 As Range, _
                  区切り文字 As String, _
                  Optional 各文字で分割 As Boolean = True, _
                  Optional 空は削除 As Boolean = True)

   Dim A, B
   Dim C As Range, n As Long, i As Long, D As Long
   Dim C1(), l As Long, myA()
   D = 文字列.Count - 1
   ReDim C1(D)
   Dim mymax As Long

   For Each C In 文字列
      C1(l) = C.Value
      If 各文字で分割 = True Then
         For i = 1 To Len(区切り文字)
            C1(l) = Replace(C1(l), Mid(区切り文字, i, 1), vbTab)
         Next i
         A = Split(C1(l), vbTab)
      Else
         A = Split(C1(l), 区切り文字)
      End If
      
      If 空は削除 = True Then
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
Function SPLIT2(文字列 As String, _
                  区切り文字 As String, _
                  Optional 各文字で分割 As Boolean = True, _
                  Optional 空は削除 As Boolean = True)

   Dim A, B
   Dim n As Long, i As Long
   
   If 各文字で分割 = True Then
      For i = 1 To Len(区切り文字)
         文字列 = Replace(文字列, Mid(区切り文字, i, 1), vbTab)
      Next i
      区切り文字 = vbTab
   End If
   
   A = Split(文字列, 区切り文字)
   
   If 空は削除 = True Then
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


