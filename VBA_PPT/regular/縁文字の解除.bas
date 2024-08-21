Attribute VB_Name = "âèï∂éöÇÃâèú"
Option Explicit

Sub âèï∂éöÇÃâèú()
   Dim shp As Shape, shp2 As Shape
   Dim gcnt As Long
   On Error GoTo ErrHdl
   For Each shp In ActiveWindow.Selection.ShapeRange
      If shp.Type = msoGroup Then
         shp.Ungroup.Select
         gcnt = ActiveWindow.Selection.ShapeRange.Count
         For Each shp2 In ActiveWindow.Selection.ShapeRange
            shp2.Delete
            gcnt = gcnt - 1
            If gcnt = 1 Then Exit For
         Next shp2
      End If
   Next shp
   Exit Sub
ErrHdl:
   
End Sub

'    fcolor = Array(vbWhite, rgbBlueViolet - 500, rgbBlue + 100)
    
'    ReDim LineWidth(4)
'    ReDim fcolor(UBound(LineWidth))
'    For i = LBound(LineWidth) To UBound(LineWidth)
'        LineWidth(i) = 10 * i
''        fcolor(i) = vbBlack
''        If i Mod 2 = 1 Then
''            fcolor(i) = vbWhite
''        End If
'        fcolor(i) = RandColor
'    Next
