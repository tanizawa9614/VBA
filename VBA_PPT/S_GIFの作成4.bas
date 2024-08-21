Attribute VB_Name = "S_GIF�̍쐬4"
Option Explicit

Sub GIF�̍쐬4()
    Dim parentPrs As PowerPoint.Presentation
    Set parentPrs = ActivePresentation
    
    Dim newPrs As PowerPoint.Presentation
    Set newPrs = Application.Presentations.Add
    
    Dim newSld As PowerPoint.Slide
    Dim shp As PowerPoint.Shape
    Dim tmp As PowerPoint.PpViewType
    Dim fol As Object, f As Object
    Dim fol_path As String
    Dim titleMsg
    Dim mainTitle
    Set fol = CreateObject("Shell.Application").BrowseForFolder(0, "�摜�t�H���_�I��", &H10, 0)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If fol Is Nothing Then GoTo Fin
    
    fol_path = fol.Self.Path
    
'    If parentPrs.SlideShowWindows.Count > 0 Then parentPrs.SlideShowWindow.View.Exit
    
    With parentPrs.Windows(1)
        tmp = .ViewType
        .ViewType = ppViewSlide
    End With
    
    ' �V�K�v���[���e�[�V�����ɃX���C�h��ǉ�
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
        Set shp = newSld.Shapes.AddPicture(fileName:=fileName.Path, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=0, _
            Top:=0)
        
        If cnt = 1 Then
            ' �X���C�h�T�C�Y�̕ύX
            ratio = shp.Height / shp.Width
            With newPrs.PageSetup
                .SlideHeight = .SlideWidth * ratio
            End With
        End If
        
        ' �摜���X���C�h�T�C�Y�Ƀ��T�C�Y
        With shp
            .LockAspectRatio = True
            .Width = newPrs.PageSetup.SlideWidth
            .Height = newPrs.PageSetup.SlideHeight
        End With
        
        ' �摜���X���C�h�����ɔz�u
        With shp
            .Left = 0
            .Top = 0
        End With
        
L1:
    Next
    
Fin:
    
End Sub


