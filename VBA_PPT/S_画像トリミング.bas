Attribute VB_Name = "S_�摜�g���~���O"
Option Explicit

Sub �摜�g���~���O()
    ' �ϐ��̒�`
    Dim shps As Object
    Dim shp As Shape
    Dim a As Double
    Dim w As Double
    Dim h As Double

    Dim Sld As Slide
    Dim Si As Long
    On Error Resume Next
    Si = ActiveWindow.Selection.SlideRange(1).SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
    Set shps = ActiveWindow.Selection.ShapeRange
    If Err.Number > 0 Then Exit Sub
    On Error GoTo 0
    
    Dim r1 As Double
    Dim r2 As Double
    r1 = 0.03
    r2 = 0.2
    For Each shp In shps
        ' �I�����ꂽ�}�`�݂̂��g���~���O����
        If shp.Type = msoPicture Then
            w = shp.width
            h = shp.height
            With shp.PictureFormat
                .CropTop = h * r1
                .CropLeft = w * r2
                .CropBottom = h * r1
                .CropRight = w * r2
            End With
        End If
    Next
End Sub

