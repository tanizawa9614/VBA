Attribute VB_Name = "S_�S�摜���w��T�C�Y�ɕύX����"
Option Explicit
Sub �S�摜���w��T�C�Y�ɕύX����()

    '�ϐ��̒�`
    Dim shp As Shape
    Dim sld As Slide
    Dim a As Double

    Set sld = ActivePresentation.Slides(2)
    '�S�X���C�h����������
'    For Each sld In ActivePresentation.Slides
        '�X���C�h�ɑ��݂���S�摜����������
        For Each shp In sld.Shapes
            '�摜�̏ꍇ�̂ݏ�������
            If shp.Type = msoPicture Then
                With shp.PictureFormat
                    .CropTop = 20
                    .CropLeft = 163.6
                    .CropBottom = 20
                    .CropRight = 229.5
                End With
'                a = shp.Width
'                shp.Width = a / 13.02 * 6.47
            End If
        Next shp
'    Next sld

End Sub
Sub �}��ɑ΂���2()

    '�ϐ��̒�`
    Dim shp As Shape
    Dim sld As Slide
    Dim a As Double

    Set sld = ActivePresentation.Slides(2)
    '�S�X���C�h����������
'    For Each sld In ActivePresentation.Slides
        '�X���C�h�ɑ��݂���S�摜����������
        For Each shp In sld.Shapes
            '�摜�̏ꍇ�̂ݏ�������
            If shp.Type = msoPicture Then
                With shp.PictureFormat
                    .CropTop = 28
                    .CropLeft = 694
                    .CropBottom = 14
                    .CropRight = 20
                End With
            End If
        Next shp
'    Next sld

End Sub
Sub �S�摜���w��T�C�Y�ɕύX����v���O�������猋��()

    '�ϐ��̒�`
    Dim shp As Shape
    Dim sld As Slide
    Dim a As Double

    Set sld = ActivePresentation.Slides(2)
    '�S�X���C�h����������
'    For Each sld In ActivePresentation.Slides
        '�X���C�h�ɑ��݂���S�摜����������
        For Each shp In sld.Shapes
            '�摜�̏ꍇ�̂ݏ�������
            If shp.Type = msoPicture Then
                With shp.PictureFormat
                    .CropTop = 20
                    .CropLeft = 62
                    .CropBottom = 20
                    .CropRight = 128
                End With
'                a = shp.Width
'                shp.Width = a / 13.02 * 6.47
            End If
        Next shp
'    Next sld

End Sub
Sub �}��ɑ΂���()

    '�ϐ��̒�`
    Dim shp As Shape
    Dim sld As Slide
    Dim a As Double

    Set sld = ActivePresentation.Slides(2)
    '�S�X���C�h����������
'    For Each sld In ActivePresentation.Slides
        '�X���C�h�ɑ��݂���S�摜����������
        For Each shp In sld.Shapes
            '�摜�̏ꍇ�̂ݏ�������
            If shp.Type = msoPicture Then
                With shp.PictureFormat
                    .CropTop = 28
                    .CropLeft = 543
                    .CropBottom = 14
                    .CropRight = 20
                End With
            End If
        Next shp
'    Next sld

End Sub

