Attribute VB_Name = "F_FINDDATEANDNAME"
Option Explicit

Function FINDDATEANDNAME(A00 As Range)
    Dim A, i As Long, j As Long, A0
    Dim buf(), buf2(), tmp, buf4
    A = A00
    A = FILEDIV(A)
    A0 = A
    ReDim buf(1 To UBound(A))
    ReDim buf2(1 To UBound(A))
    ReDim buf3(1 To UBound(A))
    For i = 1 To UBound(A, 1)
        buf(i) = A(i, 2)
    Next
    For i = 1 To UBound(buf)
        For j = 1 To Len(buf(i))
            If IsNumeric(Mid(buf(i), j, 1)) Then
                buf2(i) = buf2(i) & Mid(buf(i), j, 1)
            Else
                buf3(i) = buf3(i) & Mid(buf(i), j, 1)
            End If
        Next
    Next
    For i = 1 To UBound(buf2)
        Select Case Len(buf2(i))
        Case 3, 4
            buf2(i) = Year(Now) & buf2(i)
        Case 6
            buf2(i) = "20" & buf2(i)
        Case 0
            buf2(i) = Replace(Format(Now, "yyyy/mm/dd"), "/", "")
        End Select
    Next
    tmp = Array("_", " ", "　", "＿")
    buf4 = SPLIT2(buf3, tmp)
    Dim B
    B = JOINARRAY(A0, buf2, buf4)
    FINDDATEANDNAME = B
End Function
Private Function SPLIT2(文字列, _
               区切り文字, _
               Optional 空は削除 As Boolean = True, _
               Optional 代替文字 As String = "")

   Dim A, B, i As Long, j As Long
   Dim C, n As Long, D As Long
   Dim l As Long, myA(), Delimiter, buf
   Dim mymax As Long
   
   Application.Volatile
   If IsArray(文字列) Then
      C = 文字列
      D = UBound(C, 1) - 1
   Else
      D = 0
   End If
   ReDim C1(1 To D)
   
   For Each C In 文字列
      Delimiter = 区切り文字
      If IsArray(Delimiter) Then
         For Each buf In Delimiter
            C = Replace(C, buf, vbTab)
         Next
      Else
         C = Replace(C, Delimiter, vbTab)
      End If
      A = Split(C, vbTab)
      
      If 空は削除 Then
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
   
   For i = 1 To UBound(myA, 1)
      For j = 1 To UBound(myA, 2)
         If IsEmpty(myA(i, j)) Then myA(i, j) = 代替文字
      Next
   Next

   SPLIT2 = myA
End Function

Private Function JOINARRAY(A0, A1, A2)
    Dim i As Long, A(), j As Long
    ReDim A(1 To UBound(A1), 1 To 2 + UBound(A2, 2) + 2)
    For i = 1 To UBound(A1)
        A(i, 1) = A0(i, 1)
        A(i, 2) = A1(i)
        For j = 0 To UBound(A2, 2)
            A(i, 2 + j + 1) = A2(i - 1, j)
        Next
        A(i, 2 + UBound(A2, 2) + 2) = A0(i, 3)
    Next
    JOINARRAY = A
End Function
Private Function FILEDIV(対象) As Variant
   Dim o, i As Long, myA()
   ReDim myA(2, UBound(対象) - 1)
   Application.Volatile
   For Each o In 対象
      myA(0, i) = Left(o, InStrRev(o, "\"))
      If InStrRev(o, ".") - InStrRev(o, "\") > 0 Then
         myA(1, i) = Mid(o, InStrRev(o, "\") + 1, InStrRev(o, ".") - InStrRev(o, "\") - 1)
      Else
         myA(1, i) = Mid(o, InStrRev(o, "\") + 1)
      End If
      If InStrRev(o, ".") - InStrRev(o, "\") > 0 Then
         myA(2, i) = Right(o, Len(o) - InStrRev(o, ".") + 1)
      Else
         myA(2, i) = """フォルダです"""
      End If
      i = i + 1
   Next o
   FILEDIV = WorksheetFunction.Transpose(myA)
End Function
