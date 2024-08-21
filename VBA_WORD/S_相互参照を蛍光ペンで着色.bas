Attribute VB_Name = "S_相互参照を蛍光ペンで着色"
Option Explicit

Sub 相互参照を蛍光ペンで着色()

  'フィールドコードを表示　※1
  ActiveWindow.View.ShowFieldCodes = True

  '蛍光ペンの色を緑色に指定
  Options.DefaultHighlightColorIndex = wdYellow
'  Options.DefaultHighlightColorIndex = wdNone

  '蛍光ペンで着色する文字の条件を指定
  With Selection.Find

    '検索の文字列と書式のクリア
    .ClearFormatting

    '置き換えの文字列と書式のクリア
    .Replacement.ClearFormatting

    '置換語の文字列の蛍光ペンをオン
    .Replacement.Highlight = False
    .Replacement.Font.Color = wdColorRed

    'フィールドコードを対象文字に設定
    .Text = "^d"
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchWildcards = False
    .MatchFuzzy = False

    '蛍光ペンで着色を実行　※2
    .Execute Replace:=wdReplaceAll

  End With

  'フィールドコードを非表示　※3
  ActiveWindow.View.ShowFieldCodes = False

End Sub
