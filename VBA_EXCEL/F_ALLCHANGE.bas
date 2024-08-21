Attribute VB_Name = "F_ALLCHANGE"
Option Explicit
Function FILEDIRCHANGE(���̃t�@�C���� As Range, �V�K�t�@�C���� As Range) As String
   Dim OF As Range, NF As Range
   Dim i As Long, j As Long
   On Error Resume Next
   If ���̃t�@�C����.Count <> �V�K�t�@�C����.Count Then
      FILEDIRCHANGE = "�Q�Ƃ��s���ł�"
      Exit Function
   End If
   For Each OF In ���̃t�@�C����
      j = j + 1
      i = 0
      For Each NF In �V�K�t�@�C����
         i = i + 1
         If i = j Then
            Name OF As NF
            Exit For
         End If
      Next NF
   Next OF
   FILEDIRCHANGE = "�������܂���!"
End Function

Function ALLCHANGE(�Ώ� As Range) As Variant
   Dim C As Range, myA(), i As Long
   ReDim myA(�Ώ�.Count - 1)
   On Error Resume Next
   For Each C In �Ώ�
      If InStr(C, "���T��R�i��") >= 1 Then
         myA(i) = DETECTIVE(C)
      ElseIf InStr(C, "�΃D�����邷�܂�") >= 1 Then
         myA(i) = SELLS(C)
      Else
         myA(i) = C
      End If
      i = i + 1
   Next C
   ALLCHANGE = WorksheetFunction.Transpose(myA)
End Function
Function SELLS(�Ώ� As Range) As Variant
   Dim A, T As String, T1 As String, T2 As String
   Dim i As Long, S1 As Long, S2 As Long
   Dim D As Long
   A = FILEDIV(�Ώ�)
   D = UBound(A) - 1
   i = 1
   On Error Resume Next
   T = A(2)
   S1 = InStr(T, "�u")
   S2 = InStr(T, "�v")
   T1 = Mid(T, S1 + 1, S2 - S1 - 1)
   S1 = InStr(T, "�i")
   S2 = InStr(T, "�j")
   T2 = "��" & StrConv(Mid(T, S1 + 1, S2 - S1 - 1), vbNarrow) & "�b "
   If T2 = "" Then T2 = "���ʕ� "
   T = A(1) & T2 & T1 & A(3)
   SELLS = T
End Function

Function DETECTIVE(�Ώ� As Range) As Variant
   Dim A, T As String, T1 As String, T2 As String
   Dim i As Long, S1 As Long, S2 As Long
   Dim D As Long
   A = FILEDIV(�Ώ�)
   D = UBound(A, 1) - 1
  
   i = 1
   T = A(2)
   S1 = InStr(T, "�u")
   S2 = InStr(T, "�v")
   T1 = Mid(T, S1, S2 - S1 + 1)
   S1 = InStr(T, "��")
   S2 = InStr(T, "�b")
   T2 = Mid(T, S1, S2 - S1 + 1)
   T = A(1) & T2 & " " & T1 & A(3)
   DETECTIVE = T
End Function
Function FILEDIV(�Ώ� As Range) As Variant
   Dim o As Range, i As Long, myA()
   ReDim myA(2, �Ώ�.Count - 1)
   For Each o In �Ώ�
      myA(0, i) = Left(o, InStrRev(o, "\"))
      myA(1, i) = Mid(o, InStrRev(o, "\") + 1, InStrRev(o, ".") - InStrRev(o, "\") - 1)
      myA(2, i) = Right(o, Len(o) - InStrRev(o, ".") + 1)
      i = i + 1
   Next o
   FILEDIV = WorksheetFunction.Transpose(myA)
End Function
