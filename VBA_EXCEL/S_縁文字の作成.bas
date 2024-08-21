Attribute VB_Name = "S_縁文字の作成"
Option Explicit

Sub 縁文字の作成()
Attribute 縁文字の作成.VB_ProcData.VB_Invoke_Func = " \n14"
'作成日時：2023/03/14
'選択中の図形に対して全て縁文字に変更します
'nshapeで縁文字を何重にするかを設定できる
'LineWidthで縁文字の太さを変更できる
'縁文字の文字色についてはダイヤログボックスから変更できる

    On Error GoTo ErrHandl
    Application.ScreenUpdating = False
    Application.EnableCancelKey = xlErrorHandler
    
    Dim nshape As Long, i As Long, j As Long
    Dim shp As Shape, sname()
    Dim shp2 As Shape
    Dim LineWidth, fcolor()
    
    nshape = 3 ' 何重の文字にするか
    LineWidth = Array(4, 7) '上から順に線の太さを指定，大きさは nshape-1
    nshape = UBound(LineWidth) + 2
    
    ReDim sname(nshape - 1), fcolor(nshape - 2)
    
    ' カラーダイヤログボックスの表示
    For i = 0 To nshape - 2
        Application.Dialogs(xlDialogEditColor).Show (1)
        fcolor(i) = ActiveWorkbook.Colors(1)
    Next
    
    ' 選択中の図形に対して実行
    For Each shp In Selection.ShapeRange
        sname(0) = shp.Name
        
        ' 複製・輪郭の太さを指定，色の設定
        For i = 1 To nshape - 1
            Set shp2 = shp.Duplicate
            sname(i) = shp2.Name
            With shp2.TextFrame2.TextRange.Font.Line
                .Visible = msoTrue
                .Weight = LineWidth(i - 1)
                .ForeColor.RGB = fcolor(i - 1)
            End With
        Next
        
        ' 複製した図形を複数選択
        For i = 0 To nshape - 1
            If i = 0 Then
                ActiveSheet.Shapes(sname(i)).Select '図形を選択
            Else
                ActiveSheet.Shapes(sname(i)).Select Replace:=False '図形を「追加]
            End If
        Next
        
        ' 上下左右中央揃え・グループ化
        Selection.ShapeRange.Align msoAlignMiddles, msoFalse
        Selection.ShapeRange.Align msoAlignCenters, msoFalse
        Selection.ShapeRange.Group.Select
        
        ' 図形の並び替え
        For i = 1 To nshape - 1
            ActiveSheet.Shapes(sname(i)).Select
            For j = 1 To i
                Selection.ShapeRange.ZOrder msoSendBackward
            Next
        Next
    Next
    
Exit Sub
ErrHandl:
    MsgBox "Error です"
    Application.ScreenUpdating = True
  
End Sub
