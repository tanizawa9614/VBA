Attribute VB_Name = "S_全画像を指定サイズと"
Option Explicit

Sub 全画像を指定サイズと指定位置に変更する()

    '変数の定義
    Dim shp As Shape
    Dim sld As Slide

    '全スライドを処理する
    For Each sld In ActivePresentation.Slides
        'スライドに存在する全画像を処理する
        For Each shp In sld.Shapes
            '画像の場合のみ処理する
            If shp.Type = msoPicture Then
        
                '縦横比を固定するのチェックをはずす
                shp.LockAspectRatio = msoFalse
                '画像のサイズを変更する
                shp.Width = 72 / 2.54 * 10 '横幅を10cmにする
                shp.Height = 72 / 2.54 * 10 '縦幅を10cmにする
                '画像を指定座標に移動する
                shp.Left = 72 / 2.54 * 2 'X座標を2cmにする
                shp.Top = 72 / 2.54 * 2 'Y座標を2cmにする
            End If
        Next shp
    Next sld

End Sub
