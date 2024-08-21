Attribute VB_Name = "S_GIF�̍쐬"
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
    
    '�X���C�h�T�C�Y�̕ύX
    With ActivePresentation.PageSetup
        .SlideWidth = 150
        .SlideHeight = 150
    End With
    
    '�t�H���_���̃t�@�C������
    With CreateObject("Scripting.FileSystemObject")
        If Not .FolderExists(fol_path) Then GoTo Fin
        For Each f In .GetFolder(fol_path).Files
            'JPEG�t�@�C���̂ݏ���
            Select Case LCase(.GetExtensionName(f.Path))
                Case "jpg", "jpeg", "png"
                Set sld = prs.Slides.Add(prs.Slides.Count + 1, ppLayoutBlank)
                sld.Select
                ' �摜�ǉ�
                Set shp = sld.Shapes.AddPicture(FileName:=f.Path, _
                LinkToFile:=False, _
                SaveWithDocument:=True, _
                Left:=0, _
                Top:=0)
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
            End Select
        Next
    End With
    
'    ActivePresentation.Slides.Range.Export FileName:=fol_path & "\animation.gif", filtername:="gif"
Fin:
    ActiveWindow.ViewType = tmp
    
    
End Sub

