Attribute VB_Name = "S_�_���p�t�H�[�}�b�g����"
Option Explicit
Sub �_���p�t�H�[�}�b�g����()
Attribute �_���p�t�H�[�}�b�g����.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'***1.�����̒u���C2.�t�H���g�̒u��
'***3.�󔒍폜,4.�擪�y�[�W�Ɉړ�
   Dim i As Long, A As String
   Const Font1 = "�l�r ����"
   Const Font2 = "Times New Roman"
   
'�����̒u��
'���_�G���p�u.�v�ɑ΂��Ēu���������Ȃ��Ƃ�����
   Dim H_before, H_after
   H_before = �u���Ώ�
   H_after = �u����̕���
   For i = 0 To UBound(H_before)
      With Selection.Find
         .Text = H_before(i)
         .Replacement.Text = H_after(i)
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .MatchCase = True
         .MatchWholeWord = True
         .MatchByte = True
         .MatchAllWordForms = False
         .MatchSoundsLike = False
         .MatchWildcards = False
         .MatchFuzzy = False
         .Execute Replace:=wdReplaceAll
      End With
   Next i
'�t�H���g�̒u��

   With Selection
      .WholeStory
      .Font.Name = Font1
      .Font.Name = Font2
   End With
   With ActiveDocument
      For i = 1 To .Paragraphs.Count
         With .Paragraphs(i).Range
            A = ActiveDocument.Range(.Start, .End).Style
            If A <> "�W��" Then _
            ActiveDocument.Range(.Start, .End).Style = A
         End With
      Next
   End With

'�󔒂̍폜
'���_�GSelection �����炵����,
'        �p�������X�^�C���̐ݒ��Font2
'�@      �ł͂Ȃ��ꍇ�X�y�[�X�͏�������
   Dim pcnt As Long, j As Long, buf As String
   Dim s_start As Long, s_stop As Long
'   On Error Resume Next
   pcnt = ActiveDocument.Content.Information(4)
   For j = 1 To pcnt
      Application.ScreenUpdating = False
      s_start = s_stop
      If j = pcnt Then
         s_stop = ActiveDocument.Characters.Count
      Else
         s_stop = Selection.GoTo(What:=wdGoToPage, _
                  Which:=wdGoToAbsolute, _
                  Count:=j + 1).End - 1
      End If
      Application.ScreenUpdating = True
      Selection.GoTo What:=wdGoToPage, _
                  Which:=wdGoToAbsolute, _
                  Count:=j
      buf = MsgBox(j & "�y�[�W�ڂ��������܂����H", vbYesNo)
      If buf = vbYes Then
         On Error GoTo L1
         For i = s_start To s_stop
            If i > s_stop Then Exit For
            If ActiveDocument.Range(i, i + 1) = " " Then
               If _
               ActiveDocument.Range(i + 1, i + 2).Font.Name <> Font2 Or _
               ActiveDocument.Range(i - 1, i).Font.Name <> Font2 Then
                  ActiveDocument.Range(i, i + 1) = ""
                  s_stop = s_stop - 1
               End If
            ElseIf ActiveDocument.Range(i, i + 1) = "�@" Then
               ActiveDocument.Range(i, i + 1) = ""
               s_stop = s_stop - 1
            End If
         Next i
      End If
   Next j
L1:
'�擪�y�[�W�Ɉړ�
   Selection.GoTo What:=wdGoToPage, _
                  Which:=wdGoToAbsolute, _
                  Count:=1
   MsgBox "�I�����܂���"
End Sub

Function �u���Ώ�()
   Dim A(5)
   A(0) = "�B"
   A(1) = "."
   A(2) = "�A"
   A(3) = ","
   A(4) = "("
   A(5) = ")"
   �u���Ώ� = A
End Function
Function �u����̕���()
   Dim A(5)
   A(0) = "�D"
   A(1) = "�D"
   A(2) = "�C"
   A(3) = "�C"
   A(4) = "�i"
   A(5) = "�j"
   �u����̕��� = A
End Function
