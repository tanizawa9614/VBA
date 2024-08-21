Attribute VB_Name = "Module1"
Option Explicit
Function TEXTAFTER _
(•¶š—ñ As Range, ‹æØ‚è•¶š As String, Optional ˆÊ’u As Long = 1, Optional •¶ší‚Å‹æ•Ê As Boolean = True)
   Dim A, B As String, myA(1), i As Long, C(6) As String, D As String
   If •¶ší‚Å‹æ•Ê Then
      A = Split(•¶š—ñ, ‹æØ‚è•¶š)
   Else
      C(0) = StrConv(‹æØ‚è•¶š, vbUpperCase)
      C(1) = StrConv(‹æØ‚è•¶š, vbLowerCase)
      C(2) = StrConv(‹æØ‚è•¶š, vbProperCase)
      C(3) = StrConv(‹æØ‚è•¶š, vbKatakana)
      C(4) = StrConv(C(3), vbWide)
      C(5) = StrConv(C(3), vbNarrow)
      C(6) = StrConv(‹æØ‚è•¶š, vbHiragana)
      D = •¶š—ñ
      For i = 0 To 6
         D = Replace(D, C(i), vbTab)
      Next i
      A = Split(D, vbTab)
   End If
   
   If UBound(A) + 1 <= ˆÊ’u Then
      TEXTAFTER = Err.Description
      Exit Function
   End If
   
   For i = 0 To ˆÊ’u - 1
      B = B & A(i) & Mid(•¶š—ñ, Len(B & A(i)) + 1, Len(‹æØ‚è•¶š))
   Next i
   myA(0) = B
   myA(1) = Mid(•¶š—ñ, Len(B) + 1, Len(•¶š—ñ))
   TEXTAFTER = myA(1)
End Function
