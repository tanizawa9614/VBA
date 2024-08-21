Attribute VB_Name = "三重ふち文字の作成"
Option Explicit
Sub 三重ふち文字の作成()
   Dim D As Slide, Si As Long, i As Long, shp As Shape
   Si = ActiveWindow.Selection.SlideRange.SlideIndex
   Set D = ActivePresentation.Slides(Si)
   Dim Col As Variant, L As Long, T As Long, LiW As Long
   Dim font_size As Double, N As String
   Dim ratio As Double, FL As Double, SL As Double
   ratio = 0.035 '元の文字の大きさ
   FL = 0.37     '2個目の輪郭
   SL = 0.52     '3個目の輪郭
   
   On Error Resume Next
   For Each shp In ActiveWindow.Selection.ShapeRange
      With shp
         If .Type = msoGroup Then GoTo L1
         L = .Left
         T = .Top
         font_size = .TextFrame2.TextRange.Font.size
         LiW = Round(ratio * font_size, 1)
         Col = .TextFrame.TextRange.Font.Color.RGB
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
                     .Weight = Round(ratio ^ -FL * LiW, 1)
                     .ForeColor.RGB = RGB(255, 255, 255)
                  End With
               Else
                  With .TextFrame2.TextRange.Font.Line
                     .Weight = Round(ratio ^ -SL * LiW, 1)
                     .ForeColor.RGB = Col
                  End With
                  .ZOrder msoSendBackward
               End If
               .ZOrder msoSendBackward
            End With
         Next i
         D.Shapes.Range(Array(N & "1", N & "2", N & "3")).Group.Name = N
      End With
L1:
   Next shp
End Sub

