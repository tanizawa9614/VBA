Attribute VB_Name = "S_10���J��"
Option Explicit

Sub S_10���J��()
   Dim Target_t As Range, Target_R As Range
   Dim write_log As Range
   On Error GoTo myErr
   Const msg = "�I������͈͂͏c�����̒P���łȂ���΂Ȃ�܂���" & vbCr
   Set Target_t = Application.InputBox _
         ("�^�����B���� t �ɊY������Z���͈͂�I�����Ă�������" _
         & vbCr & vbCr & "��)" & msg, Type:=8, Title:="�y�Z���͈͂̑I�� 1/3�z �^�����B���� t")
   Set Target_R = Application.InputBox _
         ("�^�����B���ԉJ�� R �ɊY������Z���͈͂�I�����Ă�������" _
         & vbCr & vbCr & "��)" & msg, Type:=8, Title:="�y�Z���͈͂̑I�� 2/3�z �^�����B���ԉJ�� R")
   Set write_log = Application.InputBox _
         ("���ʂ̏o�͐�Z����I�����Ă�������" & vbCr, _
         Type:=8, Title:="�y�Z���͈͂̑I�� 3/3�z ���ʂ̏o�͐�")
   Dim t, R
   t = Target_t.Value
   R = Target_R.Value
   If UBound(t, 1) = 1 Or UBound(R, 1) = 1 Or UBound(t, 2) >= 2 Or UBound(R, 2) >= 2 Then
'      t = WorksheetFunction.Transpose(t) '�����炭���̊֐��͎g���Ȃ��̂�
'      R = WorksheetFunction.Transpose(R)
      MsgBox "t�����R��" & msg
      End
   End If
   
   Dim n As Long, i As Long, j As Long
   Dim t2(), t2_10(), sta_n(), end_n()
   Dim row_sum As Long

   n = UBound(t, 1)
   ReDim t2(n - 1), t2_10(n - 1)
   ReDim sta_n(n - 1), end_n(n - 1)

   sta_n(0) = 0
   For i = 0 To n - 1
      t2(i) = t(i + 1, 1) - sta_n(i)
      t2_10(i) = Int(t2(i) / 10)
      end_n(i) = t2(i) Mod 10
      If i <> n - 1 Then sta_n(i + 1) = 10 - end_n(i)
      row_sum = row_sum + t2_10(i)
      If end_n(i) <> 0 Then row_sum = row_sum + 1
   Next i
   
   Dim buf(), cnt As Long, row_count As Long
   ReDim buf(row_sum - 1, 3 * n - 1)
   For j = 0 To 3 * n - 1
      If ((j + 1) + 1) Mod 3 = 0 Then
         For cnt = 1 To t2_10(Int(j / 3))
            buf(row_count, j) = 10
            row_count = row_count + 1
         Next
      ElseIf ((j + 1) + 1) Mod 3 = 2 Then
         buf(row_count, j) = sta_n(Int(j / 3))
         If sta_n(Int(j / 3)) <> 0 Then row_count = row_count + 1
      Else
         If end_n(Int(j / 3)) = 0 Then row_count = row_count - 1
         buf(row_count, j) = end_n(Int(j / 3))
      End If
   Next
   Dim ans()
   ReDim ans(row_sum - 1, 0)
   For i = 0 To UBound(buf, 1)
      For j = 0 To UBound(buf, 2)
         ans(i, 0) = ans(i, 0) + buf(i, j) * R(Int(j / 3) + 1, 1) / t(Int(j / 3) + 1, 1)
      Next
   Next
myErr:
   Const Errmsg = vbCr & "�@ �Z���͈͂�I�����Ȃ������D" & vbCr _
   & "�A �Z���ȊO�̂��̂���͂���" & vbCr & "�B ���̑�(�J���҂֑��k)"
   If Err.Number > 0 Then
      MsgBox "�G���[���������܂���.�l�����錴���͈ȉ��̒ʂ�ł��D" & Errmsg
      Exit Sub
   End If
   Dim flag
   If WorksheetFunction.CountA(write_log.Resize(row_sum)) >= 1 Then
      flag = MsgBox("���ʂ��o�͂��悤�Ƃ��Ă���Z���͈͓̔�(����)�Ɋ����̃f�[�^�����݂��Ă��܂��D" _
      & vbCr & "���ʂ��o�͂��܂����H", vbYesNo + vbQuestion)
      If flag = vbNo Then
         MsgBox "�����𒆒f���܂���"
         Exit Sub
      End If
   End If
'   write_log.Resize(row_sum, 3 * n).Offset(, 1) = buf
   write_log.Resize(row_sum) = ans
End Sub
