Attribute VB_Name = "F_FILEDIRCHANGE"
Option Explicit

Function FILEDIRCHANGE(元のファイル名 As Range, 新規ファイル名 As Range) As String
   Dim OF As Range, NF As Range
   Dim i As Long, j As Long, flag As String
   On Error Resume Next
   Application.Volatile
   If 元のファイル名.Count <> 新規ファイル名.Count Then
      FILEDIRCHANGE = "参照が不正です"
      Exit Function
   End If
   flag = MsgBox("ファイル名の変更を行いますか？", vbQuestion + vbYesNo)
   If flag = vbNo Then
      FILEDIRCHANGE = "処理が中止されました"
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

Function FILEDIV(対象 As Range) As Variant
   Dim o As Range, i As Long, myA()
   ReDim myA(2, 対象.Count - 1)
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
