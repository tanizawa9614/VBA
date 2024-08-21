Attribute VB_Name = "S_縁文字の作成および解除"
Option Explicit

Sub 縁文字の作成()  'ppt用
Attribute 縁文字の作成.VB_ProcData.VB_Invoke_Func = " \n14"
'作成日時：2023/03/14
'選択中の図形に対して全て縁文字に変更します
'LineWidthで縁文字の太さを変更できる
'fcolorで縁文字の色を指定
'縁文字の文字色についてはダイヤログボックスから変更できる

    On Error GoTo ErrHandl
 
    Dim nshape As Long, i As Long, j As Long
    Dim shp As Shape, sname()
    Dim LineWidth(), fcolor()
    
    Dim Sld As Slide, Si As Long
    Si = ActiveWindow.Selection.SlideRange.SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
    
    LineWidth = Array(6, 8, 10, 14) '上から順に線の太さを指定，大きさは nshape-1
    fcolor = Array(vbWhite, RandColor, vbWhite, RandColor)
    
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
    
    If UBound(LineWidth) <> UBound(fcolor) Then
        MsgBox "LineWidth配列とfcolor配列の大きさを同じにしてください"
        Exit Sub
    End If
    
    nshape = UBound(LineWidth) + 2
    
    ReDim sname(nshape - 1)
    
    ' 選択中の図形に対して実行
    For Each shp In ActiveWindow.Selection.ShapeRange
        sname(0) = shp.Name
        
        ' 複製・輪郭の太さを指定，色の設定
        For i = 1 To nshape - 1
            With shp.Duplicate
                sname(i) = .Name
                With .TextFrame2.TextRange.Font.Line
                    .Visible = msoTrue
                    .Weight = LineWidth(i - 1)
                    .ForeColor.RGB = fcolor(i - 1)
                End With
            End With
        Next
        
        ' 複製した図形を複数選択
        For i = 0 To nshape - 1
            If i = 0 Then
                Sld.Shapes(sname(i)).Select '図形を選択
            Else
                Sld.Shapes(sname(i)).Select Replace:=False '図形を「追加]
            End If
        Next
        
        ' 上下左右中央揃え・グループ化
        With ActiveWindow.Selection.ShapeRange
            .Align msoAlignMiddles, msoFalse
            .Align msoAlignCenters, msoFalse
            .Group.Select
        End With
        
        ' 図形の並び替え
        For i = 1 To nshape - 1
            For j = 1 To i
                Sld.Shapes(sname(i)).ZOrder msoSendBackward
            Next
        Next
    Next
    
Exit Sub

ErrHandl:
    MsgBox "Error です"
      
End Sub

Private Function RandColor() As Long
    Randomize
    Dim minN As Long, maxN As Long
    Dim r As Long, g As Long, b As Long
    minN = 0
    maxN = 255
    r = Int((maxN - minN + 1) * Rnd + minN)
    g = Int((maxN - minN + 1) * Rnd + minN)
    b = Int((maxN - minN + 1) * Rnd + minN)
    RandColor = RGB(r, g, b)
'    RandColor = Array(r, g, b)
End Function

Sub 縁文字の解除()
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

