Attribute VB_Name = "重文字の作成・解除_ゼミ用"
Option Explicit
Sub 呼び出し_三重ふち文字の作成()
   Call 三重ふち文字の作成
End Sub
Sub 呼び出し_三重文字の解除()
   Call 三重文字の解除
End Sub

Sub 三重ふち文字の作成()
   Dim D As Slide, Si As Long, i As Long, shp As Shape
   Si = ActiveWindow.Selection.SlideRange.SlideIndex
   Set D = ActivePresentation.Slides(Si)
   Dim Col As Long, Col2 As Long, Col3 As Long
   Dim B_Col As Variant
   Dim L As Long, T As Long, LiW As Long
   Dim font_size As Double, N As String
   Dim ratio As Double, FL As Double, SL As Double
   
   Col = vbWhite
   Col2 = 2162853
   Col3 = vbWhite
   B_Col = ""
   On Error Resume Next
   For Each shp In ActiveWindow.Selection.ShapeRange
      With shp
         If .Type = msoGroup Then GoTo L1
         L = .Left
         T = .Top
         font_size = .TextFrame2.TextRange.Font.Size
         .TextFrame.TextRange.Font.Color = Col
         LiW = 1
         B_Col = .Fill.ForeColor.RGB
         If B_Col <> "" Then
            .Fill.Visible = False
         End If
         
         With .TextFrame2.TextRange.Font.Line
            .Weight = LiW
            .ForeColor.RGB = Col
         End With
         N = .TextFrame2.TextRange.Text
         .Name = N & "1"
         .ZOrder msoBringToFront
         
         For i = 1 To 2
            With .Duplicate
               .Top = T
               .Left = L
               .Name = N & i + 1
               If i = 1 Then
                  With .TextFrame2.TextRange.Font.Line
                     .Weight = 7.3
                     .ForeColor.RGB = Col2
                  End With
               Else
                  With .TextFrame2.TextRange.Font.Line
                     .Weight = 9.3
                     .ForeColor.RGB = Col3
                  End With
                  .Fill.ForeColor.RGB = B_Col
                  .ZOrder msoSendBackward
               End If
               .ZOrder msoSendBackward
            End With
         Next i
         D.Shapes.Range(Array(N & "1", N & "2", N & "3")).Group.Name = N
      End With
L1:
   Next shp
   shp.Select
End Sub

Sub 三重文字の解除()
   Dim shp As Shape, shp2 As Shape
   Dim i As Long
   For Each shp In ActiveWindow.Selection.ShapeRange
      i = 3
      If shp.Type = msoGroup Then
         shp.Ungroup.Select
         For Each shp2 In ActiveWindow.Selection.ShapeRange
            shp2.Delete
            i = i - 1
            If i = 1 Then Exit For
         Next shp2
      End If
   Next shp
End Sub

