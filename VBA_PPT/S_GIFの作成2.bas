Attribute VB_Name = "S_GIFの作成2"
Option Explicit

Sub GIFの作成()
    Dim prs As PowerPoint.Presentation
    Dim sld As PowerPoint.Slide
    Dim shp As PowerPoint.Shape
    Dim tmp As PowerPoint.PpViewType
    Dim fol As Object, f As Object
    Dim fol_path As String
    Dim titleMsg
    Dim mainTitle
    Set prs = ActivePresentation
    Set fol = CreateObject("Shell.Application") _
    .BrowseForFolder(0, "画像フォルダ選択", &H10, 0)
    
    If fol Is Nothing Then GoTo Fin
    
    fol_path = fol.Self.Path
    
    If SlideShowWindows.Count > 0 Then prs.SlideShowWindow.View.Exit
    
    With ActiveWindow
        tmp = .ViewType
        .ViewType = ppViewSlide
    End With
    
    ' 1枚目の画像ファイルを取得してスライドサイズを設定
    With CreateObject("Scripting.FileSystemObject")
        If Not .FolderExists(fol_path) Then GoTo Fin
        For Each f In .GetFolder(fol_path).files
            Select Case LCase(.GetExtensionName(f.Path))
                Case "jpg", "jpeg", "png"
                    ' 1枚目の画像ファイルの情報を取得
                    Dim img As PictureFormat
                    Set img = prs.Slides(1).Shapes.AddPicture(fileName:=f.Path, _
                        LinkToFile:=False, _
                        SaveWithDocument:=True, _
                        Left:=0, _
                        Top:=0).PictureFormat
                    
                    ' スライドサイズの変更
                    With prs.PageSetup
                        .SlideWidth = img.Width
                        .SlideHeight = img.Height
                    End With
                    
                    ' 1枚目の画像を削除
                    prs.Slides(1).Shapes(img.Parent.Id).Delete
                    
                    Exit For
            End Select
        Next
    End With
    
    ' フォルダ内のファイル処理（名前順）
    Dim files As New Collection
    With CreateObject("Scripting.FileSystemObject")
        If Not .FolderExists(fol_path) Then GoTo Fin
        For Each f In .GetFolder(fol_path).files
            Select Case LCase(.GetExtensionName(f.Path))
                Case "jpg", "jpeg", "png"
                    files.Add f.Path
            End Select
        Next
    End With
    
    ' ファイル名順にスライドに画像を追加
    Dim fileName As Variant
    For Each fileName In files
        Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
        sld.Select
        Set shp = sld.Shapes.AddPicture(fileName:=fileName, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=0, _
            Top:=0)
        
        ' 画像リサイズ
        With shp
            .LockAspectRatio = True
            .Width = prs.PageSetup.SlideWidth
            .Height = prs.PageSetup.SlideHeight
            .Select
        End With
        
        ' 画像をスライド中央に配置
        With ActiveWindow.Selection.ShapeRange
            .Align msoAlignCenters, True
            .Align msoAlignBottoms, True
        End With
    Next
    
    ' アニメーションGIFの出力
    ActivePresentation.Slides.Range.Export fileName:=fol_path & "\animation.gif", filtername:="gif"
    
Fin:
    ActiveWindow.ViewType = tmp
    
End Sub


