Attribute VB_Name = "S_�\�p�r��"
Option Explicit

Sub �\�p�r��()
   Dim Target As Range
   Dim buf As Range
   Dim i As Long
   Dim R As Long, C As Long
   On Error GoTo L1
   Call AppActivate(ThisWorkbook.Name)
   Set Target = Application.InputBox("�ΏۃZ����I�����Ă�������", Type:=8)
   Application.ScreenUpdating = False
   
   R = Target.Rows.Count
   C = Target.Columns.Count
   Target.Borders.LineStyle = xlNone
'��[
   With Target.Resize(1)
      With .Borders(xlEdgeBottom) '��[����
'         .LineStyle = xlContinuous
         .LineStyle = xlDouble
         .Weight = xlThick
      End With
      With .Borders(xlEdgeTop) '�ŏ�[
         .LineStyle = xlContinuous
'         .LineStyle = xlDouble
         .Weight = xlThick
      End With
'���[
   End With
   With Target.Resize(1).Offset(R - 1)
      With .Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThick
      End With
   End With
'�c��
   For i = 1 To C - 1
      With Target.Resize(, 1).Offset(, i).Borders(xlEdgeLeft)
         .LineStyle = xlContinuous
      End With
   Next
'����
   For i = 2 To R - 1
      With Target.Resize(1).Offset(i).Borders(xlEdgeTop)
         .LineStyle = xlDot
      End With
   Next
L1:
   Application.ScreenUpdating = True
End Sub
