Attribute VB_Name = "F_FINDBETWEEN"
Option Explicit

Function FINDBETWEEN(�Ώ� As Range, _
                        �X�^�[�g As Variant, _
                           �X�g�b�v As Variant)
   Dim upper As Long, lower As Long
   Dim C As Range, mya(), i As Long
   ReDim mya(�Ώ�.Count - 1)
   For Each C In �Ώ�
      upper = InStr(C, �X�^�[�g)
      lower = InStr(C, �X�g�b�v)
      If upper <> 0 And lower <> 0 Then
         mya(i) = Mid(C, upper + 1, lower - upper - 1)
      Else
         mya(i) = "�Ȃ�"
      End If
      i = i + 1
   Next C
   FINDBETWEEN = WorksheetFunction.Transpose(mya)
End Function
 


