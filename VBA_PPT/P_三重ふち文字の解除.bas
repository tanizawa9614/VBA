Attribute VB_Name = "�O�d�ӂ������̉���"
Option Explicit

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
