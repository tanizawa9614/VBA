Attribute VB_Name = "S_範囲のフォントを置換"
Sub 指定文字に囲まれた範囲のフォントを置換()
    Dim rng As Range
    
    Const str0 As String = "【"  ' 指定文字の開始文字
    Const str1 As String = "】"  ' 指定文字の終了文字
    
    
    ' 文書全体を対象にするために、Rangeを文書全体に設定
    Set rng = ActiveDocument.Content
    ' デフォルトのハイライトを黄色に設定
    Options.DefaultHighlightColorIndex = wdYellow
    
    With rng.Find
        .ClearFormatting
        .MatchWildcards = True
        .Text = "\" & str0 & "*\" & str1
        .Replacement.Font.Color = wdColorBlue  ' 色を置換
        .Replacement.Highlight = True
        .Execute Replace:=wdReplaceAll ' すべて置換
    End With
    
    ' 検索を終了
    Set rng = Nothing
End Sub

