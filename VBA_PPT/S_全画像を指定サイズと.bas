Attribute VB_Name = "S_�S�摜���w��T�C�Y��"
Option Explicit

Sub �S�摜���w��T�C�Y�Ǝw��ʒu�ɕύX����()

    '�ϐ��̒�`
    Dim shp As Shape
    Dim sld As Slide

    '�S�X���C�h����������
    For Each sld In ActivePresentation.Slides
        '�X���C�h�ɑ��݂���S�摜����������
        For Each shp In sld.Shapes
            '�摜�̏ꍇ�̂ݏ�������
            If shp.Type = msoPicture Then
        
                '�c������Œ肷��̃`�F�b�N���͂���
                shp.LockAspectRatio = msoFalse
                '�摜�̃T�C�Y��ύX����
                shp.Width = 72 / 2.54 * 10 '������10cm�ɂ���
                shp.Height = 72 / 2.54 * 10 '�c����10cm�ɂ���
                '�摜���w����W�Ɉړ�����
                shp.Left = 72 / 2.54 * 2 'X���W��2cm�ɂ���
                shp.Top = 72 / 2.54 * 2 'Y���W��2cm�ɂ���
            End If
        Next shp
    Next sld

End Sub
