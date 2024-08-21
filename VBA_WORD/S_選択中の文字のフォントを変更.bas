Attribute VB_Name = "S_選択中の文字のフォントを変更"
Option Explicit

Sub 選択中の文字のフォントを変更()
    Const font1 As String = "ＭＳ ゴシック"
    Const font2 As String = "Times New Roman"
    
     
    ' テキストが選択されていることを確認
    If Selection.Type = wdSelectionNormal Then
        ' 選択したテキストのフォントを変更
        Selection.Font.NameFarEast = font1
        Selection.Font.Name = font2
    End If
End Sub

