Attribute VB_Name = "S_�������̍쐬"
Option Explicit

Sub �������̍쐬()
Attribute �������̍쐬.VB_ProcData.VB_Invoke_Func = " \n14"
'�쐬�����F2023/03/14
'�I�𒆂̐}�`�ɑ΂��đS�ĉ������ɕύX���܂�
'nshape�ŉ����������d�ɂ��邩��ݒ�ł���
'LineWidth�ŉ������̑�����ύX�ł���
'�������̕����F�ɂ��Ă̓_�C�����O�{�b�N�X����ύX�ł���

    On Error GoTo ErrHandl
    Application.ScreenUpdating = False
    Application.EnableCancelKey = xlErrorHandler
    
    Dim nshape As Long, i As Long, j As Long
    Dim shp As Shape, sname()
    Dim shp2 As Shape
    Dim LineWidth, fcolor()
    
    nshape = 3 ' ���d�̕����ɂ��邩
    LineWidth = Array(4, 7) '�ォ�珇�ɐ��̑������w��C�傫���� nshape-1
    nshape = UBound(LineWidth) + 2
    
    ReDim sname(nshape - 1), fcolor(nshape - 2)
    
    ' �J���[�_�C�����O�{�b�N�X�̕\��
    For i = 0 To nshape - 2
        Application.Dialogs(xlDialogEditColor).Show (1)
        fcolor(i) = ActiveWorkbook.Colors(1)
    Next
    
    ' �I�𒆂̐}�`�ɑ΂��Ď��s
    For Each shp In Selection.ShapeRange
        sname(0) = shp.Name
        
        ' �����E�֊s�̑������w��C�F�̐ݒ�
        For i = 1 To nshape - 1
            Set shp2 = shp.Duplicate
            sname(i) = shp2.Name
            With shp2.TextFrame2.TextRange.Font.Line
                .Visible = msoTrue
                .Weight = LineWidth(i - 1)
                .ForeColor.RGB = fcolor(i - 1)
            End With
        Next
        
        ' ���������}�`�𕡐��I��
        For i = 0 To nshape - 1
            If i = 0 Then
                ActiveSheet.Shapes(sname(i)).Select '�}�`��I��
            Else
                ActiveSheet.Shapes(sname(i)).Select Replace:=False '�}�`���u�ǉ�]
            End If
        Next
        
        ' �㉺���E���������E�O���[�v��
        Selection.ShapeRange.Align msoAlignMiddles, msoFalse
        Selection.ShapeRange.Align msoAlignCenters, msoFalse
        Selection.ShapeRange.Group.Select
        
        ' �}�`�̕��ёւ�
        For i = 1 To nshape - 1
            ActiveSheet.Shapes(sname(i)).Select
            For j = 1 To i
                Selection.ShapeRange.ZOrder msoSendBackward
            Next
        Next
    Next
    
Exit Sub
ErrHandl:
    MsgBox "Error �ł�"
    Application.ScreenUpdating = True
  
End Sub
