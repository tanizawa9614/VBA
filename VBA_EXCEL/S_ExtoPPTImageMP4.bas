Attribute VB_Name = "S_ExtoPPTImageMP4"
Option Explicit

Sub Excel����PPT�ɉ摜�t�@�C���̓\��t������ѓ��扻()
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    Dim pptShape As Object
    Dim folderPath As String
    Dim imagePath As String
    Dim imageFile As String
    Dim slideIndex As Integer
    Dim answer As Integer
    
    ' PowerPoint�A�v���P�[�V�������J��
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    
    ' �V�����v���[���e�[�V�������쐬
    Set pptPres = pptApp.Presentations.Add
    
    ' �X���C�h�̃T�C�Y��ݒ�i4:3�j
    pptPres.PageSetup.SlideWidth = 914.4 ' 10�C���`
    pptPres.PageSetup.SlideHeight = 685.8 ' 7.5�C���`
    
    ' �摜�t�H���_�̑I��
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "�摜�t�H���_��I�����Ă�������"
        .Show
        If .SelectedItems.Count > 0 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "�t�H���_���I������Ă��܂���B�������I�����܂��B", vbExclamation
            Exit Sub
        End If
    End With
    
    ' �摜�t�H���_���̃t�@�C�����擾
    imageFile = Dir(folderPath & "\*.*") ' ���ׂĂ̊g���q�ɑΉ�
    
    Do While imageFile <> ""
               
        If LCase(Right(imageFile, 4)) Like ".jpg" Or LCase(Right(imageFile, 4)) Like ".jpeg" Or _
           LCase(Right(imageFile, 4)) Like ".png" Or LCase(Right(imageFile, 4)) Like ".gif" Then
            ' �X���C�h��ǉ�
            slideIndex = pptPres.Slides.Count + 1
            Set pptSlide = pptPres.Slides.Add(slideIndex, 12) ' ppLayoutBlank �̒l: 12
            
            ' �摜���X���C�h�ɓ\��t��
            imagePath = folderPath & "\" & imageFile
            Set pptShape = pptSlide.Shapes.AddPicture(Filename:=imagePath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=0, Top:=0)
            pptShape.LockAspectRatio = msoTrue
            pptShape.Top = (pptPres.PageSetup.SlideHeight - pptShape.Height) / 2
            pptShape.Left = (pptPres.PageSetup.SlideWidth - pptShape.Width) / 2
            pptShape.ScaleWidth 1, msoFalse ' �摜�̕����X���C�h�̕��ɍ��킹��
            pptShape.ScaleHeight 1, msoFalse ' �摜�̍������X���C�h�̍����ɍ��킹��
        End If
        
        ' ���̉摜�t�@�C�����擾
        imageFile = Dir
    Loop
        
    ' MP4�ւ̕ϊ����m�F
    answer = MsgBox("MP4�ɕϊ����܂����H", vbYesNo + vbQuestion)
    If answer = vbYes Then
        ' �v���[���e�[�V������MP4�Ƃ��ĕۑ�
        pptPres.SaveAs folderPath & "\output.mp4", 39 ' ppSaveAsMP4 �̒l: 39
        pptPres.Close
        MsgBox "MP4�t�@�C���Ƃ��ďo�͂��܂����B�������I�����܂��B", vbInformation
    Else
        ' PowerPoint��\�����ďI��
        pptApp.Visible = True
        MsgBox "�������I�����܂��B", vbInformation
    End If
    
    ' �I�u�W�F�N�g�����
    Set pptShape = Nothing
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
End Sub

