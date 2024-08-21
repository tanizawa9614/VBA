Attribute VB_Name = "F_ALLCHANGE"
Option Explicit
Function FILEDIRCHANGE(元のファイル名 As Range, 新規ファイル名 As Range) As String
   Dim OF As Range, NF As Range
   Dim i As Long, j As Long
   On Error Resume Next
   If 元のファイル名.Count <> 新規ファイル名.Count Then
      FILEDIRCHANGE = "参照が不正です"
      Exit Function
   End If
   For Each OF In 元のファイル名
      j = j + 1
      i = 0
      For Each NF In 新規ファイル名
         i = i + 1
         If i = j Then
            Name OF As NF
            Exit For
         End If
      Next NF
   Next OF
   FILEDIRCHANGE = "完了しました!"
End Function

Function ALLCHANGE(対象 As Range) As Variant
   Dim C As Range, myA(), i As Long
   ReDim myA(対象.Count - 1)
   On Error Resume Next
   For Each C In 対象
      If InStr(C, "名探偵コナン") >= 1 Then
         myA(i) = DETECTIVE(C)
      ElseIf InStr(C, "笑ゥせぇるすまん") >= 1 Then
         myA(i) = SELLS(C)
      Else
         myA(i) = C
      End If
      i = i + 1
   Next C
   ALLCHANGE = WorksheetFunction.Transpose(myA)
End Function
Function SELLS(対象 As Range) As Variant
   Dim A, T As String, T1 As String, T2 As String
   Dim i As Long, S1 As Long, S2 As Long
   Dim D As Long
   A = FILEDIV(対象)
   D = UBound(A) - 1
   i = 1
   On Error Resume Next
   T = A(2)
   S1 = InStr(T, "「")
   S2 = InStr(T, "」")
   T1 = Mid(T, S1 + 1, S2 - S1 - 1)
   S1 = InStr(T, "（")
   S2 = InStr(T, "）")
   T2 = "第" & StrConv(Mid(T, S1 + 1, S2 - S1 - 1), vbNarrow) & "話 "
   If T2 = "" Then T2 = "特別編 "
   T = A(1) & T2 & T1 & A(3)
   SELLS = T
End Function

Function DETECTIVE(対象 As Range) As Variant
   Dim A, T As String, T1 As String, T2 As String
   Dim i As Long, S1 As Long, S2 As Long
   Dim D As Long
   A = FILEDIV(対象)
   D = UBound(A, 1) - 1
  
   i = 1
   T = A(2)
   S1 = InStr(T, "「")
   S2 = InStr(T, "」")
   T1 = Mid(T, S1, S2 - S1 + 1)
   S1 = InStr(T, "第")
   S2 = InStr(T, "話")
   T2 = Mid(T, S1, S2 - S1 + 1)
   T = A(1) & T2 & " " & T1 & A(3)
   DETECTIVE = T
End Function
Function FILEDIV(対象 As Range) As Variant
   Dim o As Range, i As Long, myA()
   ReDim myA(2, 対象.Count - 1)
   For Each o In 対象
      myA(0, i) = Left(o, InStrRev(o, "\"))
      myA(1, i) = Mid(o, InStrRev(o, "\") + 1, InStrRev(o, ".") - InStrRev(o, "\") - 1)
      myA(2, i) = Right(o, Len(o) - InStrRev(o, ".") + 1)
      i = i + 1
   Next o
   FILEDIV = WorksheetFunction.Transpose(myA)
End Function
