Attribute VB_Name = "S_GIFの作成3"
Option Explicit

Sub GIFの作成3()
    Dim newPrs As PowerPoint.Presentation
    Dim newSld As PowerPoint.Slide
    Dim shp As PowerPoint.Shape
    Dim tmp As PowerPoint.PpViewType
    Dim fol As Object, f As Object
    Dim fol_path As String
    Dim titleMsg
    Dim mainTitle
    Set fol = CreateObject("Shell.Application") _
    .BrowseForFolder(0, "画像フォルダ選択", &H10, 0)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If fol Is Nothing Then GoTo Fin
    
    fol_path = fol.Self.Path
    
    If SlideShowWindows.Count > 0 Then ActivePresentation.SlideShowWindow.View.Exit
    
    With ActiveWindow
        tmp = .ViewType
        .ViewType = ppViewSlide
    End With
    
    ' 新規プレゼンテーションを作成
    Set newPrs = Presentations.Add
'    Set newSld = newPrs.Slides.Add(1, ppLayoutBlank)
    
    
    
    ' ファイル名順にスライドに画像を追加
    Dim fileName As Object
    Dim ratio As Double, cnt As Long
    For Each fileName In FSO.GetFolder(fol_path).files
        Select Case LCase(FSO.GetExtensionName(fileName.Path))
                Case "jpg", "jpeg", "png"
                Case Else
                    GoTo L1
            End Select
        cnt = cnt + 1
        Set newSld = newPrs.Slides.Add(cnt, ppLayoutBlank)
        newSld.Select
        Set shp = newSld.Shapes.AddPicture(fileName:=fileName, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=0, _
            Top:=0)
        
        If cnt = 1 Then
            ' スライドサイズの変更
            ratio = shp.Height / shp.Width
            With newPrs.PageSetup
                .SlideHeight = .SlideWidth * ratio
            End With
        End If
        ' 画像リサイズ
        With shp
            .LockAspectRatio = True
            .Width = newPrs.PageSetup.SlideWidth
            .Height = newPrs.PageSetup.SlideHeight
            .Select
        End With
        
        ' 画像をスライド中央に配置
        With ActiveWindow.Selection.ShapeRange
            .Align msoAlignCenters, True
            .Align msoAlignBottoms, True
        End With
L1:
    Next
    
    ' アニメーションGIFの出力
'    newPrs.Slides.Range.Export fileName:=fol_path & "\animation.gif", filtername:="gif"
    
Fin:
    
End Sub


' ' フォルダ内のファイル処理（名前順）
'    Dim files As New Collection
'    With FSO
'        If Not .FolderExists(fol_path) Then GoTo Fin
'        For Each f In .GetFolder(fol_path).files
'            Select Case LCase(.GetExtensionName(f.Path))
'                Case "jpg", "jpeg", "png"
'                    files.Add f.Path
'            End Select
'        Next
'    End With
