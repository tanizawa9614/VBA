Attribute VB_Name = "�d�����̍쐬�E����_�F�w��\"
Option Explicit
Sub �Ăяo��_�O�d�ӂ������̍쐬()
   Call �O�d�ӂ������̍쐬
End Sub
Sub �Ăяo��_�O�d�����̉���()
   Call �O�d�����̉���
End Sub

Sub �O�d�ӂ������̍쐬()
   Dim D As Slide, Si As Long, i As Long, shp As Shape
   Si = ActiveWindow.Selection.SlideRange.SlideIndex
   Set D = ActivePresentation.Slides(Si)
   Dim Col As Long, Col2 As Long, Col3 As Long
   Dim B_Col As Variant
   Dim L As Long, T As Long, LiW As Long
   Dim font_size As Double, N As String
   Dim ratio As Double, FL As Double, SL As Double
   ratio = 0.03 '���̕����̑傫��
   FL = 0.5     '2�ڂ̗֊s
   SL = 0.7     '3�ڂ̗֊s
   Col = vbBlack
   Col2 = vbWhite
   Col3 = vbBlue
   B_Col = ""
   On Error Resume Next
   For Each shp In ActiveWindow.Selection.ShapeRange
      With shp
         If .Type = msoGroup Then GoTo L1
         L = .Left
         T = .Top
         font_size = .TextFrame2.TextRange.Font.Size
         .TextFrame.TextRange.Font.Color = Col
         LiW = Round(ratio * font_size, 1)
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
                     .Weight = Round(ratio ^ -FL * LiW, 1)
                     .ForeColor.RGB = Col2
                  End With
               Else
                  With .TextFrame2.TextRange.Font.Line
                     .Weight = Round(ratio ^ -SL * LiW, 1)
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

Sub �O�d�����̉���()
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

