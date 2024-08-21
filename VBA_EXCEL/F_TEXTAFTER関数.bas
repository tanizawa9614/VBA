Attribute VB_Name = "Module1"
Option Explicit
Function TEXTAFTER _
(������ As Range, ��؂蕶�� As String, Optional �ʒu As Long = 1, Optional ������ŋ�� As Boolean = True)
   Dim A, B As String, myA(1), i As Long, C(6) As String, D As String
   If ������ŋ�� Then
      A = Split(������, ��؂蕶��)
   Else
      C(0) = StrConv(��؂蕶��, vbUpperCase)
      C(1) = StrConv(��؂蕶��, vbLowerCase)
      C(2) = StrConv(��؂蕶��, vbProperCase)
      C(3) = StrConv(��؂蕶��, vbKatakana)
      C(4) = StrConv(C(3), vbWide)
      C(5) = StrConv(C(3), vbNarrow)
      C(6) = StrConv(��؂蕶��, vbHiragana)
      D = ������
      For i = 0 To 6
         D = Replace(D, C(i), vbTab)
      Next i
      A = Split(D, vbTab)
   End If
   
   If UBound(A) + 1 <= �ʒu Then
      TEXTAFTER = Err.Description
      Exit Function
   End If
   
   For i = 0 To �ʒu - 1
      B = B & A(i) & Mid(������, Len(B & A(i)) + 1, Len(��؂蕶��))
   Next i
   myA(0) = B
   myA(1) = Mid(������, Len(B) + 1, Len(������))
   TEXTAFTER = myA(1)
End Function
