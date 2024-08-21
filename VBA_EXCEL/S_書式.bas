Attribute VB_Name = "S_����"
Option Explicit

Private Sub macro1()
   Call ����_�\�p�r��
End Sub
Private Sub macro2()
   Call ����_�O���t�̏����ݒ�
End Sub

Sub ����_�\�p�r��()
   Dim Target As Range
   Dim buf As Range
   Dim i As Long
   Dim R As Long, C As Long
   On Error Resume Next
   Call AppActivate(ThisWorkbook.Name)
'   Set Target = Application.InputBox("�ΏۃZ����I�����Ă�������", Type:=8)
   Set Target = Selection
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

Sub ����_�O���t�̏����ݒ�()
   Dim i As Long, j As Long
   Dim ax As Object, SC As Object
   Dim White_Color, Black_Color
   Dim axColor, buf As String
   Dim ax_max(1), ax_min(1)
   
   Const Font1 = "Times New Roman"
   Const Font2 = "�l�r �S�V�b�N"
   Const FontSize = 12
   Const Marker_Size = 5.8
   Const axLabel_Size = FontSize + 3
   
   i = 100 '�ڐ����̐F��ݒ�C�����قǖ��邢(255�܂�)
   axColor = RGB(i, i, i)
   
   White_Color = RGB(255, 255, 255)
   Black_Color = RGB(0, 0, 0)
   
   For i = 1 To ActiveSheet.ChartObjects.Count
      With ActiveSheet.ChartObjects(i).Chart
'�^�C�g���̐ݒ�
         .HasTitle = False
'�t�H���g�̐ݒ�,�O���t�G���A�g���̍폜
         With .ChartArea.Format
            .Line.Visible = msoFalse
            With .TextFrame2.TextRange.Font
               .NameComplexScript = Font1
               .NameFarEast = Font2
               .Name = Font1
               .Size = FontSize
               .Fill.ForeColor.RGB = Black_Color
            End With
         End With
'�}��̏����ݒ�
         .HasLegend = True
         With .Legend
            .IncludeInLayout = False
            .Format.Fill.ForeColor.RGB = White_Color
            .Format.Line.ForeColor.RGB = Black_Color
            With .Format.Shadow
               .Type = msoShadow21
               .Size = 102
            End With
         End With
'�n��̏����ݒ�
         j = 0
         For Each SC In .SeriesCollection
            SC.MarkerStyle = Array(8, 3, 1, 2)(j Mod 4)
            If UBound(SC.Values) >= 10 ^ 2 Then SC.MarkerStyle = xlNone
            SC.Format.Line.DashStyle = Array(1, 4, 5, 6, 7, 8, 2)(j Mod 7)
            SC.MarkerSize = Marker_Size
            SC.Format.Line.ForeColor.RGB = Black_Color
            SC.Format.Fill.ForeColor.RGB = Black_Color
            j = j + 1
         Next
'���̐ݒ�
         j = 0
         For Each ax In .Axes
            If Not ax.HasTitle Then
               ax.HasTitle = True  '���^�C�g���̍쐬
               Select Case ax.Type
                  Case 1
                     buf = "x"
                  Case 2
                     buf = "y"
               End Select
               ax.AxisTitle.Text = _
                  InputBox("��" & ax.Type & "���i" _
                  & buf & "���j�̎����x������͂��Ă�������")
            End If
            ax.AxisTitle.Font.Size = axLabel_Size
            ax.Format.Fill.ForeColor.RGB = White_Color
            ax.Format.Line.ForeColor.RGB = axColor
            ax.MajorGridlines.Format.Line.ForeColor.RGB = axColor
         Next
      End With
   Next
End Sub

Sub ����t�H���g�ɕύX()
   With Selection.Font
      .Name = "�l�r �S�V�b�N"
      .Name = "Times New Roman"
   End With
End Sub

