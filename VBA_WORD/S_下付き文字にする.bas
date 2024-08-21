Attribute VB_Name = "S_下付き文字にする"
Option Explicit

Sub 下付き文字にする()
    Dim doc As Document
    Dim rng As Range
    Dim findText As String
    Dim n As Long
    
    ' 文書オブジェクトを取得
    Set doc = ActiveDocument
    
    ' 検索対象のテキストを設定
    findText = "SC-CO2"
    ' 上記文字の何文字目を変更するか
    n = 6
    
    ' 文書内のテキストを検索し、見つかった場合に下付き文字に変更
    For Each rng In doc.StoryRanges
        With rng.Find
            .ClearFormatting
            .Text = findText
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute
            
            Do While .Found
'                rng.Select
                ' 下付き文字に変更
                rng.Characters(n).Font.Subscript = True
                ' 次の検索へ進む
                rng.Collapse wdCollapseEnd
                .Execute
            Loop
        End With
    Next rng
End Sub

