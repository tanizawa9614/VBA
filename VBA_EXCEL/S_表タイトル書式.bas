Attribute VB_Name = "S_�\�^�C�g������"
Option Explicit
   Dim i As Long, buf As String
   Dim j As Long, k As Long
   Dim St As Long, Sp As Long
   Dim Blank_place As Long
   Dim C As Range
Sub �\�^�C�g������()
   '�����@�@[���O] [�L��][�i�P�ʁj]�̌`�ɂȂ��Ă��邱��
   '�����A�@[���O]��[�L��]�̊Ԃɂ͕K�����p�X�y�[�X�����邱��
   '******�@[�P��]�͖����Ă��悢
   On Error Resume Next
   Application.ScreenUpdating = True
   Selection.Value = Selection.Value
   For Each C In Selection
      buf = C.Value
      Call ���݂̏���������
      St = InStr(buf, "[")
      Sp = InStr(buf, "]")
      Blank_place = InStr(buf, " ")
      If Mid(buf, St - 1, 1) <> " " Then
         Call ����L���̑O�ɋ󔒂�ǉ�
      End If
      If St <> 0 Then Call �P�ʂ���t����
      If Blank_place <> 0 Then
         Call �L�����Α̂�
         If Mid(buf, Blank_place + 2, 1) <> " " Then _
         Call �L���̓񕶎��ڈȍ~�����t����
      End If
      St = 0
      Sp = 0
      Blank_place = 0
   Next
   With Selection.Font
      .Name = "�l�r �S�V�b�N"
      .Name = "Times New Roman"
      .Color.RGB = RGB(0, 0, 0)
   End With
   Application.ScreenUpdating = True
End Sub
Private Sub ���݂̏���������()
   C = buf
   With C.Characters(1, Len(C)).Font
      .Italic = False
      .Subscript = False
      .Superscript = False
   End With
End Sub
Private Sub ����L���̑O�ɋ󔒂�ǉ�()
   C = Left(buf, St - 1) & " " & Mid(buf, St)
   St = St + 1
   Sp = Sp + 1
   buf = C
End Sub
Private Sub �P�ʂ���t����()
   For i = St To Sp
      If IsNumeric(Mid(buf, i, 1)) Then
         C.Characters(i, 1).Font.Superscript = True
      End If
   Next
End Sub
Private Sub �L�����Α̂�()
   C.Characters(Blank_place + 1, 1).Font.Italic = True
End Sub
Private Sub �L���̓񕶎��ڈȍ~�����t����()
   C.Characters(Blank_place + 2, _
   InStrRev(buf, " ") - Blank_place - 2).Font.Subscript = True
End Sub
