Attribute VB_Name = "S_書式_グラフ書式"
Option Explicit
Sub 書式_グラフの書式設定()
   Dim i As Long, j As Long
   Dim ax As Object, SC As Object
   Dim White_Color, Black_Color
   Dim axColor, buf As String
   
   Const Font1 = "Times New Roman"
   Const Font2 = "ＭＳ ゴシック"
   Const FontSize = 12
   Const Marker_Size = 5.8
   Const axLabel_Size = FontSize + 2
   
   i = 100 '目盛線の色を設定，高いほど明るい(255まで)
   axColor = RGB(i, i, i)
   
   White_Color = RGB(255, 255, 255)
   Black_Color = RGB(0, 0, 0)
   
   For i = 1 To ActiveSheet.ChartObjects.Count
      With ActiveSheet.ChartObjects(i).Chart
'タイトルの設定
         .HasTitle = False
'フォントの設定,グラフエリア枠線の削除
         With .ChartArea.Format
            .Line.Visible = msoFalse
            With .TextFrame2.TextRange.Font
               .NameComplexScript = Font1
               .NameFarEast = Font2
               .Name = Font1
               .Size = FontSize
               .Fill.ForeColor.RGB = Black_Color
            End With
         End With
'凡例の書式設定
         .HasLegend = True
         With .Legend
            .IncludeInLayout = False
            .Format.Fill.ForeColor.RGB = White_Color
            .Format.Line.ForeColor.RGB = Black_Color
            With .Format.Shadow
               .Type = msoShadow21
               .Size = 102
            End With
         End With
'系列の書式設定
         j = 0
         For Each SC In .SeriesCollection
            SC.MarkerStyle = Array(8, 3, 1, 2)(j Mod 4)
            SC.Format.Line.DashStyle = Array(1, 4, 5, 6, 7, 8, 2)(j Mod 7)
            SC.MarkerSize = Marker_Size
            SC.Format.Line.ForeColor.RGB = Black_Color
            SC.Format.Fill.ForeColor.RGB = Black_Color
            j = j + 1
         Next
'軸の設定
         For Each ax In .Axes
            If Not ax.HasTitle Then
               ax.HasTitle = True  '軸タイトルの作成
               Select Case ax.Type
                  Case 1
                     buf = "x"
                  Case 2
                     buf = "y"
               End Select
               ax.AxisTitle.Text = _
                  InputBox("第" & ax.Type & "軸（" _
                  & buf & "軸）の軸ラベルを入力してください")
            End If
            ax.AxisTitle.Font.Size = axLabel_Size
            ax.Format.Fill.ForeColor.RGB = White_Color
            ax.Format.Line.ForeColor.RGB = axColor
            ax.MajorGridlines.Format.Line.ForeColor.RGB = axColor
         Next
      End With
   Next
End Sub

