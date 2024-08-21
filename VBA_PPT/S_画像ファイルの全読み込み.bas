Attribute VB_Name = "S_�摜�t�@�C���̑S�ǂݍ���"
Option Explicit

Sub �摜�t�@�C���̑S�ǂݍ���_�����`()

' ���̃}�N���́C�w�肵���t�H���_���ɑ��݂���摜�t�@�C�����ꖇ���t�@�C�������ɓǂݍ��݁C
'�@�ꖇ���̃X���C�h���쐬����}�N���ł��D


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
    .BrowseForFolder(0, "�摜�t�H���_�I��", &H10, 0)
    
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
    
    '�X���C�h�T�C�Y�̕ύX
    With ActivePresentation.PageSetup
        .SlideWidth = 150
        .SlideHeight = 150
    End With
    
    cnt = 1
    
    '�t�H���_���̃t�@�C������
    With CreateObject("Scripting.FileSystemObject")
        If Not .FolderExists(fol_path) Then GoTo Fin
        For Each f In .GetFolder(fol_path).Files
            'JPEG�t�@�C���̂ݏ���
            Select Case LCase(.GetExtensionName(f.Path))
                Case "jpg", "jpeg", "png"
                If cnt = 1 Then
                    page_cnt = prs.Slides.Count '������Ԃ̃y�[�W�𐔂��Ă���
                End If
                Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
                sld.Select
                          
                ' �摜�ǉ�
                Set shp = sld.Shapes.AddPicture(FileName:=f.Path, LinkToFile:=False, SaveWithDocument:=True, Left:=0, Top:=0)
                '�摜���T�C�Y
                With shp
                    .LockAspectRatio = True
                    .Width = prs.PageSetup.SlideWidth
                    .Height = prs.PageSetup.SlideHeight
                    .Select
                End With
                '�摜���X���C�h�����ɔz�u
                With ActiveWindow.Selection.ShapeRange
                    .Align msoAlignCenters, True
                    .Align msoAlignBottoms, True
                End With
                cnt = cnt + 1
                
                
                '==================================================================================================================
                '�l�ƒP�ʂ��\�L����ꍇ
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
                        .NameFarEast = "�l�r �S�V�b�N"
                        .Line.Weight = 1.8
                        .Line.ForeColor.RGB = RGB(255, 255, 255)
                    End With
                    
                     With shp3.TextFrame2.TextRange.Font
                        .size = 8
                        .Fill.ForeColor.RGB = RGB(0, 0, 0)
                        .Name = "Arial"
                        .NameFarEast = "�l�r �S�V�b�N"
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
    If page_cnt = 1 Then '�}�N���J�n�O��������ԂȂ�1�y�[�W�ڂ͍폜
        ActivePresentation.Slides(1).Delete
    End If
    
Fin:
    ActiveWindow.ViewType = tmp
    
    
End Sub

