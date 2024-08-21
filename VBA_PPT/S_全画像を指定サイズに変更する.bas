Attribute VB_Name = "S_全画像を指定サイズに変更する"
Option Explicit
Sub 全画像を指定サイズに変更する()

    '変数の定義
    Dim shp As Shape
    Dim sld As Slide
    Dim a As Double

    Set sld = ActivePresentation.Slides(2)
    '全スライドを処理する
'    For Each sld In ActivePresentation.Slides
        'スライドに存在する全画像を処理する
        For Each shp In sld.Shapes
            '画像の場合のみ処理する
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
Sub 凡例に対して2()

    '変数の定義
    Dim shp As Shape
    Dim sld As Slide
    Dim a As Double

    Set sld = ActivePresentation.Slides(2)
    '全スライドを処理する
'    For Each sld In ActivePresentation.Slides
        'スライドに存在する全画像を処理する
        For Each shp In sld.Shapes
            '画像の場合のみ処理する
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
Sub 全画像を指定サイズに変更するプログラムから結果()

    '変数の定義
    Dim shp As Shape
    Dim sld As Slide
    Dim a As Double

    Set sld = ActivePresentation.Slides(2)
    '全スライドを処理する
'    For Each sld In ActivePresentation.Slides
        'スライドに存在する全画像を処理する
        For Each shp In sld.Shapes
            '画像の場合のみ処理する
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
Sub 凡例に対して()

    '変数の定義
    Dim shp As Shape
    Dim sld As Slide
    Dim a As Double

    Set sld = ActivePresentation.Slides(2)
    '全スライドを処理する
'    For Each sld In ActivePresentation.Slides
        'スライドに存在する全画像を処理する
        For Each shp In sld.Shapes
            '画像の場合のみ処理する
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

