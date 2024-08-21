Attribute VB_Name = "Module1"
Option Explicit
Function DETECTIVE(ëŒè€ As Range) As Variant
   Dim A, T(), T1(), T2()
   Dim i As Long, S1 As Long, S2 As Long
   Dim D As Long
   A = FILEDIV(ëŒè€)
   D = UBound(A, 1) - 1
   ReDim T(D)
   ReDim T1(D)
   ReDim T2(D)
   
   For i = 1 To UBound(A, 1)
      T(i - 1) = A(i, 2)
      S1 = InStr(T(i - 1), "Åu")
      S2 = InStr(T(i - 1), "Åv")
      T1(i - 1) = Mid(T(i - 1), S1, S2 - S1 + 1)
      S1 = InStr(T(i - 1), "ëÊ")
      S2 = InStr(T(i - 1), "òb")
      T2(i - 1) = Mid(T(i - 1), S1, S2 - S1 + 1)
      T(i - 1) = A(i, 1) & T2(i - 1) & " " & T1(i - 1) & A(i, 3)
   Next i
   DETECTIVE = WorksheetFunction.Transpose(T)
End Function
Function FILEDIV(ëŒè€ As Range) As Variant
   Dim A, B As String
   Dim O As Range
   Dim i As Long, j As Long
   Dim myA()
   ReDim myA(2, ëŒè€.Count - 1)
   Const C = "\"
   Const C2 = "."
   
   For Each O In ëŒè€
      A = Split(O, C)
      B = ""
      For i = 0 To UBound(A) - 1
         B = B & A(i) & C
      Next i
      myA(0, j) = B
      B = ""
      A = Split(A(UBound(A)), C2)
      For i = 0 To UBound(A) - 1
         If i <> UBound(A) - 1 Then
            B = B & A(i) & C2
         Else
            B = B & A(i)
         End If
      Next i
      myA(1, j) = B
      myA(2, j) = C2 & A(UBound(A))
      j = j + 1
   Next O
   FILEDIV = WorksheetFunction.Transpose(myA)
End Function
