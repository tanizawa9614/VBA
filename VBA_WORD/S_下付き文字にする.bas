Attribute VB_Name = "S_���t�������ɂ���"
Option Explicit

Sub ���t�������ɂ���()
    Dim doc As Document
    Dim rng As Range
    Dim findText As String
    Dim n As Long
    
    ' �����I�u�W�F�N�g���擾
    Set doc = ActiveDocument
    
    ' �����Ώۂ̃e�L�X�g��ݒ�
    findText = "SC-CO2"
    ' ��L�����̉������ڂ�ύX���邩
    n = 6
    
    ' �������̃e�L�X�g���������A���������ꍇ�ɉ��t�������ɕύX
    For Each rng In doc.StoryRanges
        With rng.Find
            .ClearFormatting
            .Text = findText
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute
            
            Do While .Found
'                rng.Select
                ' ���t�������ɕύX
                rng.Characters(n).Font.Subscript = True
                ' ���̌����֐i��
                rng.Collapse wdCollapseEnd
                .Execute
            Loop
        End With
    Next rng
End Sub

