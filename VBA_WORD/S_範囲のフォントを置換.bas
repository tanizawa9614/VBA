Attribute VB_Name = "S_�͈͂̃t�H���g��u��"
Sub �w�蕶���Ɉ͂܂ꂽ�͈͂̃t�H���g��u��()
    Dim rng As Range
    
    Const str0 As String = "�y"  ' �w�蕶���̊J�n����
    Const str1 As String = "�z"  ' �w�蕶���̏I������
    
    
    ' �����S�̂�Ώۂɂ��邽�߂ɁARange�𕶏��S�̂ɐݒ�
    Set rng = ActiveDocument.Content
    ' �f�t�H���g�̃n�C���C�g�����F�ɐݒ�
    Options.DefaultHighlightColorIndex = wdYellow
    
    With rng.Find
        .ClearFormatting
        .MatchWildcards = True
        .Text = "\" & str0 & "*\" & str1
        .Replacement.Font.Color = wdColorBlue  ' �F��u��
        .Replacement.Highlight = True
        .Execute Replace:=wdReplaceAll ' ���ׂĒu��
    End With
    
    ' �������I��
    Set rng = Nothing
End Sub

