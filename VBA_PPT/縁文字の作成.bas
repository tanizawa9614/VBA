Attribute VB_Name = "�������̍쐬"
Option Explicit
Dim LineWidth(), fcolor()

Sub �������̍쐬()  'ppt�p
    
    'User �ݒ�ӏ��@***************************************************************************************************
    LineWidth = Array(7, 15) '       �����̗֊s���̑����̐ݒ�,�����_OK
    fcolor = Array(vbWhite, RandColor) '�֊s���̐F�̎w��CRandColor�Ɠ��͂���ƃ����_���ȐF��Ԃ�
    '���F�͌��F�Ȃ� �@�@�@�@[vb]�̂��Ƃ�black��yellow�Ȃǁ@ �@��jvbBlack , vbBlue, vbYellow, vbRed, vbCyan
    '�������������G�ȐF�Ȃ� [rgb]�̂��Ƃ�black��yellow�Ȃǁ@�@��jrgbAliceBlue, rgbBrown, rgbDarkGreen
    '��RGB�l���w�肷��Ȃ�@[RGB(*,*,*)] *:0~255�܂ł̐����l�@��jRGB(0,0,0):�� , RGB(255,255,255):�� , RGB(255,0,0):��
    '******************************************************************************************************************
    
    'Log
    '�쐬�����F2023/03/14
    '�C�������F2023/03/21 �������쐬��Ɉʒu���ω����Ȃ��悤�ɕύX���܂���
    '�I�𒆂̐}�`�ɑ΂��đS�ĉ������ɕύX���܂�
    'LineWidth�ŉ������̑�����ύX�ł���
    'fcolor�ŉ������̐F���w��
    '�������̕����F�ɂ��Ă̓_�C�����O�{�b�N�X����ύX�ł���

    On Error GoTo ErrHandl
 
    Dim nshape As Long, i As Long, j As Long
    Dim shp As Shape, sname()
    Dim T As Double, L As Double
    
    Dim Sld As Slide, Si As Long
    Si = ActiveWindow.Selection.SlideRange.SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
    
    If UBound(LineWidth) <> UBound(fcolor) Then
        MsgBox "LineWidth�z���fcolor�z��̑傫���𓯂��ɂ��Ă�������"
        Exit Sub
    End If
    
    nshape = UBound(LineWidth) + 2
    
    ReDim sname(nshape - 1)
    
    ' �I�𒆂̐}�`�ɑ΂��Ď��s
    For Each shp In ActiveWindow.Selection.ShapeRange
        sname(0) = shp.Name
        T = shp.Top
        L = shp.Left
        If shp.Type = msoGroup Then
            GoTo Continue
'            shp.Ungroup
        End If
'        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

        ' �����E�֊s�̑������w��C�F�̐ݒ�
        For i = 1 To nshape - 1
            With shp.Duplicate
                sname(i) = .Name
                .Top = T
                .Left = L
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
            .Group.Name = shp.TextFrame2.TextRange.Text
            .Select
        End With
        
        ' �}�`�̕��ёւ�
        For i = 1 To nshape - 1
            For j = 1 To i
                Sld.Shapes(sname(i)).ZOrder msoSendBackward
            Next
        Next
Continue:
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
'    RandColor = Int((16777215 - 0 + 1) * Rnd + 0)
'    RandColor = Array(r, g, b)
End Function


