Attribute VB_Name = "S_�������̍쐬����щ���"
Option Explicit

Sub �������̍쐬()  'ppt�p
Attribute �������̍쐬.VB_ProcData.VB_Invoke_Func = " \n14"
'�쐬�����F2023/03/14
'�I�𒆂̐}�`�ɑ΂��đS�ĉ������ɕύX���܂�
'LineWidth�ŉ������̑�����ύX�ł���
'fcolor�ŉ������̐F���w��
'�������̕����F�ɂ��Ă̓_�C�����O�{�b�N�X����ύX�ł���

    On Error GoTo ErrHandl
 
    Dim nshape As Long, i As Long, j As Long
    Dim shp As Shape, sname()
    Dim LineWidth(), fcolor()
    
    Dim Sld As Slide, Si As Long
    Si = ActiveWindow.Selection.SlideRange.SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
    
    LineWidth = Array(6, 8, 10, 14) '�ォ�珇�ɐ��̑������w��C�傫���� nshape-1
    fcolor = Array(vbWhite, RandColor, vbWhite, RandColor)
    
'    ReDim LineWidth(4)
'    ReDim fcolor(UBound(LineWidth))
'    For i = LBound(LineWidth) To UBound(LineWidth)
'        LineWidth(i) = 10 * i
''        fcolor(i) = vbBlack
''        If i Mod 2 = 1 Then
''            fcolor(i) = vbWhite
''        End If
'        fcolor(i) = RandColor
'    Next
    
    If UBound(LineWidth) <> UBound(fcolor) Then
        MsgBox "LineWidth�z���fcolor�z��̑傫���𓯂��ɂ��Ă�������"
        Exit Sub
    End If
    
    nshape = UBound(LineWidth) + 2
    
    ReDim sname(nshape - 1)
    
    ' �I�𒆂̐}�`�ɑ΂��Ď��s
    For Each shp In ActiveWindow.Selection.ShapeRange
        sname(0) = shp.Name
        
        ' �����E�֊s�̑������w��C�F�̐ݒ�
        For i = 1 To nshape - 1
            With shp.Duplicate
                sname(i) = .Name
                With .TextFrame2.TextRange.Font.Line
                    .Visible = msoTrue
                    .Weight = LineWidth(i - 1)
                    .ForeColor.RGB = fcolor(i - 1)
                End With
            End With
        Next
        
        ' ���������}�`�𕡐��I��
        For i = 0 To nshape - 1
            If i = 0 Then
                Sld.Shapes(sname(i)).Select '�}�`��I��
            Else
                Sld.Shapes(sname(i)).Select Replace:=False '�}�`���u�ǉ�]
            End If
        Next
        
        ' �㉺���E���������E�O���[�v��
        With ActiveWindow.Selection.ShapeRange
            .Align msoAlignMiddles, msoFalse
            .Align msoAlignCenters, msoFalse
            .Group.Select
        End With
        
        ' �}�`�̕��ёւ�
        For i = 1 To nshape - 1
            For j = 1 To i
                Sld.Shapes(sname(i)).ZOrder msoSendBackward
            Next
        Next
    Next
    
Exit Sub

ErrHandl:
    MsgBox "Error �ł�"
      
End Sub

Private Function RandColor() As Long
    Randomize
    Dim minN As Long, maxN As Long
    Dim r As Long, g As Long, b As Long
    minN = 0
    maxN = 255
    r = Int((maxN - minN + 1) * Rnd + minN)
    g = Int((maxN - minN + 1) * Rnd + minN)
    b = Int((maxN - minN + 1) * Rnd + minN)
    RandColor = RGB(r, g, b)
'    RandColor = Array(r, g, b)
End Function

Sub �������̉���()
   Dim shp As Shape, shp2 As Shape
   Dim gcnt As Long
   For Each shp In ActiveWindow.Selection.ShapeRange
      If shp.Type = msoGroup Then
         shp.Ungroup.Select
         gcnt = ActiveWindow.Selection.ShapeRange.Count
         For Each shp2 In ActiveWindow.Selection.ShapeRange
            shp2.Delete
            gcnt = gcnt - 1
            If gcnt = 1 Then Exit For
         Next shp2
      End If
   Next shp
End Sub

