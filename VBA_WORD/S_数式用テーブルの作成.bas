Attribute VB_Name = "S_数式用テーブルの作成"
Option Explicit

Sub 数式用テーブルの作成()
    
    Dim tbl As Table
    Dim bord As border
    Dim cl As Cell

    ' 段落のスタイルを標準に設定
    Selection.ParagraphFormat.Style = ActiveDocument.Styles("標準")

    ' 新しいテーブルを挿入
    Set tbl = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=1, NumColumns:=3, _
                     DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed)
    
    ' テーブルスタイルを設定
    tbl.Style = "表 (格子)"
    
    With tbl.Range
        .Font.Name = "ＭＳ 明朝"
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        ' セルの配置と罫線を設定
        For Each cl In .Cells
            cl.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            cl.VerticalAlignment = wdCellAlignVerticalCenter
            cl.Height = 30
            For Each bord In cl.Borders
                bord.LineStyle = wdLineStyleNone
            Next
        Next
    End With
    
    tbl.Cell(2, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
        
End Sub


