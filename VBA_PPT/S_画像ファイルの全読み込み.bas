Attribute VB_Name = "S_画像ファイルの全読み込み"
Option Explicit

Sub 画像ファイルの全読み込み_正方形()

' このマクロは，指定したフォルダ内に存在する画像ファイルを一枚ずつファイル名順に読み込み，
'　一枚ずつのスライドを作成するマクロです．


    Dim prs As PowerPoint.Presentation
    Dim sld As PowerPoint.Slide
    Dim shp As PowerPoint.Shape
    Dim tmp As PowerPoint.PpViewType
    Dim fol As Object, f As Object
    Dim fol_path As String
    Dim titleMsg, cnt As Long, page_cnt As Long
    Dim mainTitle
    Set prs = ActivePresentation
    Set fol = CreateObject("Shell.Application") _
    .BrowseForFolder(0, "画像フォルダ選択", &H10, 0)
    
    Dim unit As String, PVal As Double, print_Val As Boolean, splitname As Variant
    Dim tmpname As String, printstring As String
    
    print_Val = True
    unit = "MPa"
    
    
    If fol Is Nothing Then GoTo Fin
        
    fol_path = fol.Self.Path
    
    If SlideShowWindows.Count > 0 Then prs.SlideShowWindow.View.Exit
    
    With ActiveWindow
        tmp = .ViewType
        .ViewType = ppViewSlide
    End With
    
    'スライドサイズの変更
    With ActivePresentation.PageSetup
        .SlideWidth = 150
        .SlideHeight = 150
    End With
    
    cnt = 1
    
    'フォルダ内のファイル処理
    With CreateObject("Scripting.FileSystemObject")
        If Not .FolderExists(fol_path) Then GoTo Fin
        For Each f In .GetFolder(fol_path).Files
            'JPEGファイルのみ処理
            Select Case LCase(.GetExtensionName(f.Path))
                Case "jpg", "jpeg", "png"
                If cnt = 1 Then
                    page_cnt = prs.Slides.Count '初期状態のページを数えておく
                End If
                Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
                sld.Select
                          
                ' 画像追加
                Set shp = sld.Shapes.AddPicture(FileName:=f.Path, LinkToFile:=False, SaveWithDocument:=True, Left:=0, Top:=0)
                '画像リサイズ
                With shp
                    .LockAspectRatio = True
                    .Width = prs.PageSetup.SlideWidth
                    .Height = prs.PageSetup.SlideHeight
                    .Select
                End With
                '画像をスライド中央に配置
                With ActiveWindow.Selection.ShapeRange
                    .Align msoAlignCenters, True
                    .Align msoAlignBottoms, True
                End With
                cnt = cnt + 1
                
                
                '==================================================================================================================
                '値と単位も表記する場合
                If print_Val Then
                    splitname = Split(f.Name, " ")
                    splitname = splitname(UBound(splitname, 1))
                    tmpname = Left(splitname, InStr(splitname, ".") - 1)
'                    tmpname = Left(splitname, InStr(splitname, unit) - 1)
                    PVal = val(tmpname)
                    tmpname = CStr(PVal)
'                    If Len(tmpname) = 2 Then
'                        tmpname = tmpname & "  "
'                    End If
'                    printstring = tmpname & " [" & unit & "]"
                    
                    Dim shp2 As Shape, shp3 As Shape, shp4 As Shape, shp5 As Shape
                                       
                    Set shp2 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 10, 10)
                    Set shp3 = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 21, 1.5, 10, 10)
                   
                    shp2.TextFrame.WordWrap = msoFalse
                    shp3.TextFrame.WordWrap = msoFalse
                    
                    With shp2.TextFrame.TextRange
                        .Text = PVal
                        .ParagraphFormat.Alignment = ppAlignLeft
                    End With
                    
                    With shp3.TextFrame.TextRange
                        .Text = "[" & unit & "]"
                        .ParagraphFormat.Alignment = ppAlignLeft
                    End With
                    
                    With shp2.TextFrame2.TextRange.Font
                        .size = 10
                        .Fill.ForeColor.RGB = RGB(0, 0, 0)
                        .Name = "Arial"
                        .NameFarEast = "ＭＳ ゴシック"
                        .Line.Weight = 1.8
                        .Line.ForeColor.RGB = RGB(255, 255, 255)
                    End With
                    
                     With shp3.TextFrame2.TextRange.Font
                        .size = 8
                        .Fill.ForeColor.RGB = RGB(0, 0, 0)
                        .Name = "Arial"
                        .NameFarEast = "ＭＳ ゴシック"
                        .Line.Weight = 1.8
                        .Line.ForeColor.RGB = RGB(255, 255, 255)
                    End With
                    
                    shp2.Duplicate
                    shp3.Duplicate
                    
                    Set shp4 = sld.Shapes(sld.Shapes.Range.Count - 1)
                    Set shp5 = sld.Shapes(sld.Shapes.Range.Count)
                    
                    shp4.Left = shp2.Left
                    shp4.Top = shp2.Top
                    shp5.Left = shp3.Left
                    shp5.Top = shp3.Top
                    
                    With shp4.TextFrame2.TextRange.Font.Line
                        .Weight = 0
                        .ForeColor.RGB = RGB(0, 0, 0)
                        .Visible = msoFalse
                    End With
                    
                    With shp5.TextFrame2.TextRange.Font.Line
                        .Weight = 0
                        .ForeColor.RGB = RGB(0, 0, 0)
                        .Visible = msoFalse
                    End With
                    
'                    Debug.Print shp3.TextFrame2.TextRange.Font.Line.Visible
                    sld.Shapes.Range(Array(shp2.Name, shp3.Name, shp4.Name, shp5.Name)).Group
                End If
            End Select
        Next
    End With
    If page_cnt = 1 Then 'マクロ開始前が初期状態なら1ページ目は削除
        ActivePresentation.Slides(1).Delete
    End If
    
Fin:
    ActiveWindow.ViewType = tmp
    
    
End Sub

