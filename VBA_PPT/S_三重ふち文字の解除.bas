Attribute VB_Name = "S_三重ふち文字の解除"
Option Explicit

Sub 三重文字の解除()
   Dim shp As Shape, shp2 As Shape
   Dim gcnt As Long
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
End Sub
