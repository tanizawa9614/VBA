Attribute VB_Name = "S_�g�����^�t�@�C����\�t��"
Option Explicit

Sub �g�����^�t�@�C����\�t��()
    On Error Resume Next
    ActiveWindow.Selection.ShapeRange(1).Copy  '�I�����Ă���}�`������ꍇ�͂�����R�s�[�C�Ȃ��ꍇ�̓R�s�[����Ă���}�`
    ActiveWindow.View.PasteSpecial DataType:=ppPasteEnhancedMetafile  '�g�����^�t�@�C���Ƃ��ē\�t��
    On Error GoTo 0
End Sub

