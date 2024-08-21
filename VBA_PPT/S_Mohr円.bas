Attribute VB_Name = "S_Mohr‰~"
Option Explicit
Option Base 1
Dim Sld As Slide, Si As Long
Dim Mag As Double

Sub Mohr‰~()
    
    Si = ActiveWindow.Selection.SlideRange.SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
    Mag = 0.3
    
    Call AddCircle(400, 0, 250)
    Call AddLine(0, 0, 800, 0, ArrowType:=True)
    Call AddLine(0, -300, 0, 300, ArrowType:=True)
    Call AddLine(150, 0, 200, 150, vbBlue)
    Call AddLine(650, 0, 200, 150, vbRed)
    Call AddLine(650, 150, 150, 150)
    Call AddLine(200, -200, 200, 200)
    Call AddPoint(600, 150)
    Call AddPoint(200, -150)
    Call AddPoint(200, 150)
    Call AddPoint(650, 0, vbRed)
    Call AddPoint(150, 0, vbBlue)
    
End Sub

Private Sub AddPoint(x As Double, y As Double, Optional Col As Long = vbBlack, Optional ObjType As Long = 1, Optional LabelAdd As Boolean = False)
    Dim r As Double
    r = 10
    With Sld.Shapes.AddShape(msoShapeOval, Mag * (x - r), Mag * (-y - r), Mag * (2 * r), Mag * (2 * r))
        If ObjType = 1 Then
            .Fill.ForeColor.RGB = Col
            .Line.Visible = msoFalse
        Else
            .Fill.Visible = msoFalse
            .Line.ForeColor.RGB = Col
        End If
    End With
    If LabelAdd Then
        With Sld.Shapes.AddLabel(msoTextOrientationHorizontal, Mag * (x - r + 10), Mag * (-y - r - 10), Mag * (2 * r), Mag * (2 * r)).TextFrame.TextRange
            .Text = "( " & x & " , " & y & " )"
            .Font.Color.RGB = Col
        End With
    End If
End Sub

Private Sub AddCircle(x As Double, y As Double, r As Double, Optional Col As Long = vbBlack, Optional ObjType As Long = 1)
    With Sld.Shapes.AddShape(msoShapeOval, Mag * (x - r), Mag * (-y - r), Mag * (2 * r), Mag * (2 * r))
        If ObjType <> 1 Then
            .Fill.ForeColor.RGB = Col
            .Line.Visible = msoFalse
        Else
            .Fill.Visible = msoFalse
            .Line.ForeColor.RGB = Col
        End If
    End With
End Sub

Private Sub AddLine(x0 As Double, y0 As Double, x1 As Double, y1 As Double, Optional Col As Long = vbBlack, Optional ArrowType As Boolean = False)
    With Sld.Shapes.AddConnector(msoConnectorStraight, Mag * (x0), Mag * (-y0), Mag * (x1), Mag * (-y1))
        .Line.ForeColor.RGB = Col
        If ArrowType Then
            .Line.EndArrowheadStyle = msoArrowheadTriangle
        End If
    End With
End Sub

Private Sub AddRectangular(L As Double, T As Double, W As Double, H As Double, Optional Col As Long = vbBlack, Optional ObjType As Long = 1)
    With Sld.Shapes.AddShape(msoShapeRectangle, Mag * (L), Mag * (T), Mag * (W), Mag * (H))
        If ObjType = 1 Then
            .Fill.ForeColor.RGB = Col
            .Line.Visible = msoFalse
        Else
            .Fill.Visible = msoFalse
            .Line.ForeColor.RGB = Col
        End If
    End With
End Sub
