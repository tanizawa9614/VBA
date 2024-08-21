Attribute VB_Name = "S_論文用フォーマット調整"
Option Explicit
Sub 論文用フォーマット調整()
Attribute 論文用フォーマット調整.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'***1.文字の置換，2.フォントの置換
'***3.空白削除,4.先頭ページに移動
   Dim i As Long, A As String
   Const Font1 = "ＭＳ 明朝"
   Const Font2 = "Times New Roman"
   
'文字の置換
'問題点；半角「.」に対して置換したくないときあり
   Dim H_before, H_after
   H_before = 置換対象
   H_after = 置換後の文字
   For i = 0 To UBound(H_before)
      With Selection.Find
         .Text = H_before(i)
         .Replacement.Text = H_after(i)
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .MatchCase = True
         .MatchWholeWord = True
         .MatchByte = True
         .MatchAllWordForms = False
         .MatchSoundsLike = False
         .MatchWildcards = False
         .MatchFuzzy = False
         .Execute Replace:=wdReplaceAll
      End With
   Next i
'フォントの置換

   With Selection
      .WholeStory
      .Font.Name = Font1
      .Font.Name = Font2
   End With
   With ActiveDocument
      For i = 1 To .Paragraphs.Count
         With .Paragraphs(i).Range
            A = ActiveDocument.Range(.Start, .End).Style
            If A <> "標準" Then _
            ActiveDocument.Range(.Start, .End).Style = A
         End With
      Next
   End With

'空白の削除
'問題点；Selection を減らしたい,
'        英数字がスタイルの設定でFont2
'　      ではない場合スペースは消失する
   Dim pcnt As Long, j As Long, buf As String
   Dim s_start As Long, s_stop As Long
'   On Error Resume Next
   pcnt = ActiveDocument.Content.Information(4)
   For j = 1 To pcnt
      Application.ScreenUpdating = False
      s_start = s_stop
      If j = pcnt Then
         s_stop = ActiveDocument.Characters.Count
      Else
         s_stop = Selection.GoTo(What:=wdGoToPage, _
                  Which:=wdGoToAbsolute, _
                  Count:=j + 1).End - 1
      End If
      Application.ScreenUpdating = True
      Selection.GoTo What:=wdGoToPage, _
                  Which:=wdGoToAbsolute, _
                  Count:=j
      buf = MsgBox(j & "ページ目を処理しますか？", vbYesNo)
      If buf = vbYes Then
         On Error GoTo L1
         For i = s_start To s_stop
            If i > s_stop Then Exit For
            If ActiveDocument.Range(i, i + 1) = " " Then
               If _
               ActiveDocument.Range(i + 1, i + 2).Font.Name <> Font2 Or _
               ActiveDocument.Range(i - 1, i).Font.Name <> Font2 Then
                  ActiveDocument.Range(i, i + 1) = ""
                  s_stop = s_stop - 1
               End If
            ElseIf ActiveDocument.Range(i, i + 1) = "　" Then
               ActiveDocument.Range(i, i + 1) = ""
               s_stop = s_stop - 1
            End If
         Next i
      End If
   Next j
L1:
'先頭ページに移動
   Selection.GoTo What:=wdGoToPage, _
                  Which:=wdGoToAbsolute, _
                  Count:=1
   MsgBox "終了しました"
End Sub

Function 置換対象()
   Dim A(5)
   A(0) = "。"
   A(1) = "."
   A(2) = "、"
   A(3) = ","
   A(4) = "("
   A(5) = ")"
   置換対象 = A
End Function
Function 置換後の文字()
   Dim A(5)
   A(0) = "．"
   A(1) = "．"
   A(2) = "，"
   A(3) = "，"
   A(4) = "（"
   A(5) = "）"
   置換後の文字 = A
End Function
