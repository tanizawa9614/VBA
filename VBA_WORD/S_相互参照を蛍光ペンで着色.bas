Attribute VB_Name = "S_���ݎQ�Ƃ��u���y���Œ��F"
Option Explicit

Sub ���ݎQ�Ƃ��u���y���Œ��F()

  '�t�B�[���h�R�[�h��\���@��1
  ActiveWindow.View.ShowFieldCodes = True

  '�u���y���̐F��ΐF�Ɏw��
  Options.DefaultHighlightColorIndex = wdYellow
'  Options.DefaultHighlightColorIndex = wdNone

  '�u���y���Œ��F���镶���̏������w��
  With Selection.Find

    '�����̕�����Ə����̃N���A
    .ClearFormatting

    '�u�������̕�����Ə����̃N���A
    .Replacement.ClearFormatting

    '�u����̕�����̌u���y�����I��
    .Replacement.Highlight = False
    .Replacement.Font.Color = wdColorRed

    '�t�B�[���h�R�[�h��Ώە����ɐݒ�
    .Text = "^d"
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchWildcards = False
    .MatchFuzzy = False

    '�u���y���Œ��F�����s�@��2
    .Execute Replace:=wdReplaceAll

  End With

  '�t�B�[���h�R�[�h���\���@��3
  ActiveWindow.View.ShowFieldCodes = False

End Sub
