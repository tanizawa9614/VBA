Attribute VB_Name = "S_GIF�̍쐬2"
Option Explicit

Sub GIF�̍쐬()
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
    .BrowseForFolder(0, "�摜�t�H���_�I��", &H10, 0)
    
    If fol Is Nothing Then GoTo Fin
    
    fol_path = fol.Self.Path
    
    If SlideShowWindows.Count > 0 Then prs.SlideShowWindow.View.Exit
    
    With ActiveWindow
        tmp = .ViewType
        .ViewType = ppViewSlide
    End With
    
    ' 1���ڂ̉摜�t�@�C�����擾���ăX���C�h�T�C�Y��ݒ�
    With CreateObject("Scripting.FileSystemObject")
        If Not .FolderExists(fol_path) Then GoTo Fin
        For Each f In .GetFolder(fol_path).files
            Select Case LCase(.GetExtensionName(f.Path))
                Case "jpg", "jpeg", "png"
                    ' 1���ڂ̉摜�t�@�C���̏����擾
                    Dim img As PictureFormat
                    Set img = prs.Slides(1).Shapes.AddPicture(fileName:=f.Path, _
                        LinkToFile:=False, _
                        SaveWithDocument:=True, _
                        Left:=0, _
                        Top:=0).PictureFormat
                    
                    ' �X���C�h�T�C�Y�̕ύX
                    With prs.PageSetup
                        .SlideWidth = img.Width
                        .SlideHeight = img.Height
                    End With
                    
                    ' 1���ڂ̉摜���폜
                    prs.Slides(1).Shapes(img.Parent.Id).Delete
                    
                    Exit For
            End Select
        Next
    End With
    
    ' �t�H���_���̃t�@�C�������i���O���j
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
    
    ' �t�@�C�������ɃX���C�h�ɉ摜��ǉ�
    Dim fileName As Variant
    For Each fileName In files
        Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
        sld.Select
        Set shp = sld.Shapes.AddPicture(fileName:=fileName, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=0, _
            Top:=0)
        
        ' �摜���T�C�Y
        With shp
            .LockAspectRatio = True
            .Width = prs.PageSetup.SlideWidth
            .Height = prs.PageSetup.SlideHeight
            .Select
        End With
        
        ' �摜���X���C�h�����ɔz�u
        With ActiveWindow.Selection.ShapeRange
            .Align msoAlignCenters, True
            .Align msoAlignBottoms, True
        End With
    Next
    
    ' �A�j���[�V����GIF�̏o��
    ActivePresentation.Slides.Range.Export fileName:=fol_path & "\animation.gif", filtername:="gif"
    
Fin:
    ActiveWindow.ViewType = tmp
    
End Sub


