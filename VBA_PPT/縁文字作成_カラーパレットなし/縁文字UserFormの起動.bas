Attribute VB_Name = "������UserForm�̋N��"
Option Explicit
Dim Sld As Slide
Public rgbVal()
Public LineWidth()
Public CancelFlg As Boolean
Public globalshp As Shape
Public Errflg As Boolean

'�W�����W���[���̃R�[�h

Sub ������UserForm�̋N��()
    Dim Si As Long
    Si = ActiveWindow.Selection.SlideRange.SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
    Dim textflg As Boolean
    Dim cnt As Long
    Dim shp As Shape
    
    cnt = 1
    CancelFlg = False
    Errflg = False
    
    textflg = False
    
    On Error Resume Next
    Set shp = ActiveWindow.Selection.ShapeRange.Item(1)
    If shp Is Nothing Then Exit Sub
    On Error GoTo 0
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Type = 17 Then
            Set globalshp = shp
            If cnt = 1 Then
                BuildUf.Show
                If CancelFlg Then Exit Sub
                cnt = cnt + 1
            End If
            Run�������̍쐬
        End If
    Next
    
End Sub

Sub Run�������̍쐬()
    
'    On Error GoTo ErrHandl
 
    Dim nshape As Long
    Dim i As Long
    Dim j As Long
    Dim sname()
    Dim T As Double, L As Double
    Dim shp As Shape
    
    Set shp = globalshp
      
    nshape = UBound(LineWidth) + 2
    
    ReDim sname(nshape - 1)
    
    
    sname(0) = shp.Name
    T = shp.Top
    L = shp.Left
    If shp.Type = msoGroup Then
        Exit Sub
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
                .ForeColor.RGB = rgbVal(i - 1)
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
    Set globalshp = ActiveWindow.Selection.ShapeRange.Item(1)
Exit Sub

ErrHandl:
      
End Sub

Sub Run�������̉���()
    Dim shp As Shape, shp2 As Shape
    Dim gcnt As Long
    On Error GoTo ErrHdl
    Set shp = globalshp
   
    If shp.Type = msoGroup Then
        shp.Ungroup.Select
        gcnt = ActiveWindow.Selection.ShapeRange.Count
        For Each shp2 In ActiveWindow.Selection.ShapeRange
            If gcnt = 1 Then Exit For
            shp2.Delete
            gcnt = gcnt - 1
        Next shp2
    Else
        Set shp2 = shp
    End If
    Set globalshp = shp2
    Exit Sub
ErrHdl:
   
End Sub


