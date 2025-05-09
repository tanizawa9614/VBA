Attribute VB_Name = "Module1"
Option Explicit
Function TEXTAFTER _
(文字列 As Range, 区切り文字 As String, Optional 位置 As Long = 1, Optional 文字種で区別 As Boolean = True)
   Dim A, B As String, myA(1), i As Long, C(6) As String, D As String
   If 文字種で区別 Then
      A = Split(文字列, 区切り文字)
   Else
      C(0) = StrConv(区切り文字, vbUpperCase)
      C(1) = StrConv(区切り文字, vbLowerCase)
      C(2) = StrConv(区切り文字, vbProperCase)
      C(3) = StrConv(区切り文字, vbKatakana)
      C(4) = StrConv(C(3), vbWide)
      C(5) = StrConv(C(3), vbNarrow)
      C(6) = StrConv(区切り文字, vbHiragana)
      D = 文字列
      For i = 0 To 6
         D = Replace(D, C(i), vbTab)
      Next i
      A = Split(D, vbTab)
   End If
   
   If UBound(A) + 1 <= 位置 Then
      TEXTAFTER = Err.Description
      Exit Function
   End If
   
   For i = 0 To 位置 - 1
      B = B & A(i) & Mid(文字列, Len(B & A(i)) + 1, Len(区切り文字))
   Next i
   myA(0) = B
   myA(1) = Mid(文字列, Len(B) + 1, Len(文字列))
   TEXTAFTER = myA(1)
End Function
