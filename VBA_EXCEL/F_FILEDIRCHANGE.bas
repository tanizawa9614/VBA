Attribute VB_Name = "F_FILEDIRCHANGE"
Option Explicit

Function FILEDIRCHANGE(���̃t�@�C���� As Range, �V�K�t�@�C���� As Range) As String
   Dim OF As Range, NF As Range
   Dim i As Long, j As Long, flag As String
   On Error Resume Next
   Application.Volatile
   If ���̃t�@�C����.Count <> �V�K�t�@�C����.Count Then
      FILEDIRCHANGE = "�Q�Ƃ��s���ł�"
      Exit Function
   End If
   flag = MsgBox("�t�@�C�����̕ύX���s���܂����H", vbQuestion + vbYesNo)
   If flag = vbNo Then
      FILEDIRCHANGE = "���������~����܂���"
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

Function FILEDIV(�Ώ� As Range) As Variant
   Dim o As Range, i As Long, myA()
   ReDim myA(2, �Ώ�.Count - 1)
   Application.Volatile
   For Each o In �Ώ�
      myA(0, i) = Left(o, InStrRev(o, "\"))
      If InStrRev(o, ".") - InStrRev(o, "\") > 0 Then
         myA(1, i) = Mid(o, InStrRev(o, "\") + 1, InStrRev(o, ".") - InStrRev(o, "\") - 1)
      Else
         myA(1, i) = Mid(o, InStrRev(o, "\") + 1)
      End If
      If InStrRev(o, ".") - InStrRev(o, "\") > 0 Then
         myA(2, i) = Right(o, Len(o) - InStrRev(o, ".") + 1)
      Else
         myA(2, i) = """�t�H���_�ł�"""
      End If
      i = i + 1
   Next o
   FILEDIV = WorksheetFunction.Transpose(myA)
End Function
