Attribute VB_Name = "S_表タイトル書式"
Option Explicit
   Dim i As Long, buf As String
   Dim j As Long, k As Long
   Dim St As Long, Sp As Long
   Dim Blank_place As Long
   Dim C As Range
Sub 表タイトル書式()
   '条件①　[名前] [記号][（単位）]の形になっていること
   '条件②　[名前]と[記号]の間には必ず半角スペースを入れること
   '******　[単位]は無くてもよい
   On Error Resume Next
   Application.ScreenUpdating = True
   Selection.Value = Selection.Value
   For Each C In Selection
      buf = C.Value
      Call 現在の書式を解除
      St = InStr(buf, "[")
      Sp = InStr(buf, "]")
      Blank_place = InStr(buf, " ")
      If Mid(buf, St - 1, 1) <> " " Then
         Call 特定記号の前に空白を追加
      End If
      If St <> 0 Then Call 単位を上付きに
      If Blank_place <> 0 Then
         Call 記号を斜体に
         If Mid(buf, Blank_place + 2, 1) <> " " Then _
         Call 記号の二文字目以降を下付きに
      End If
      St = 0
      Sp = 0
      Blank_place = 0
   Next
   With Selection.Font
      .Name = "ＭＳ ゴシック"
      .Name = "Times New Roman"
      .Color.RGB = RGB(0, 0, 0)
   End With
   Application.ScreenUpdating = True
End Sub
Private Sub 現在の書式を解除()
   C = buf
   With C.Characters(1, Len(C)).Font
      .Italic = False
      .Subscript = False
      .Superscript = False
   End With
End Sub
Private Sub 特定記号の前に空白を追加()
   C = Left(buf, St - 1) & " " & Mid(buf, St)
   St = St + 1
   Sp = Sp + 1
   buf = C
End Sub
Private Sub 単位を上付きに()
   For i = St To Sp
      If IsNumeric(Mid(buf, i, 1)) Then
         C.Characters(i, 1).Font.Superscript = True
      End If
   Next
End Sub
Private Sub 記号を斜体に()
   C.Characters(Blank_place + 1, 1).Font.Italic = True
End Sub
Private Sub 記号の二文字目以降を下付きに()
   C.Characters(Blank_place + 2, _
   InStrRev(buf, " ") - Blank_place - 2).Font.Subscript = True
End Sub
