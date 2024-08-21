Attribute VB_Name = "F_Plot"
Option Explicit
Function Plot(ParamArray data() As Variant)
    ' �v���b�g����f�[�^�̐����m�F
    Dim numData As Integer
    numData = UBound(data) - LBound(data) + 1
    Const Font1 = "Times New Roman"
    Const Font2 = "�l�r �S�V�b�N"
    Const FontSize = 12
    Const Marker_Size = 5.8
    Const axLabel_Size = FontSize + 3
    
    ' �f�[�^����2�̔{���łȂ��ꍇ�̓G���[��\�����ďI��
    If numData Mod 2 <> 0 Then
        MsgBox "�����̐����s���ł��Bx�f�[�^��y�f�[�^�̑g�ݍ��킹���w�肵�Ă��������B"
        Exit Function
    End If
    
    ' �O���t�I�u�W�F�N�g���쐬
    Dim cht As ChartObject
    Dim Xmin As Double
    Dim Xmax As Double
    Dim Ymin As Double
    Dim Ymax As Double
    
    Set cht = ActiveSheet.ChartObjects.Add(Left:=100, Width:=375, Top:=75, Height:=225)
    
    With cht.Chart
        ' �f�[�^���v���b�g
        Dim i As Long
        For i = LBound(data) To UBound(data) Step 2
            .SeriesCollection.NewSeries
            .SeriesCollection(i / 2 + 1).XValues = data(i)
            .SeriesCollection(i / 2 + 1).Values = data(i + 1)
            Xmin = MyMin(data(i), Xmin)
            Xmax = MyMax(data(i), Xmax)
            Ymin = MyMin(data(i + 1), Ymin)
            Ymax = MyMax(data(i + 1), Ymin)
        Next i
        
        ' �`���[�g��\��
'        .ChartType = xlXYScatter
'        .ChartType = xlXYScatterSmooth
        .ChartType = xlXYScatterSmoothNoMarkers
'        .ChartType = xlXYScatterLines
'        .ChartType = xlXYScatterLinesNoMarkers
        .HasTitle = False
        
        '�t�H���g�̐ݒ�,�O���t�G���A�g���̍폜
        With .ChartArea.Format
            .Line.Visible = msoFalse
            With .TextFrame2.TextRange.Font
                .NameComplexScript = Font1
                .NameFarEast = Font2
                .Name = Font1
                .Bold = msoFalse
                .Size = FontSize
                .Fill.ForeColor.RGB = rgbBlack
            End With
        End With
        
        ' �����E�c���̐ݒ�
        Dim ax As Object
        For Each ax In .Axes
            With ax
                .HasTitle = True
                .AxisTitle.Font.Size = axLabel_Size
                .Format.Fill.ForeColor.RGB = rgbWhite
                .Format.Line.ForeColor.RGB = rgbBlack
                .MajorGridlines.Format.Line.Visible = msoFalse
                .MajorTickMark = xlInside
                .MinorTickMark = xlInside
                If .Type = xlCategory Then
                    .AxisTitle.Text = "x"
                    .MinimumScale = Xmin
'                    .MaximumScale = Xmax
                    .CrossesAt = .MinimumScale
                Else
                    .AxisTitle.Text = "y"
'                    .MinimumScale = Int(Ymin)
'                    .MaximumScale = Int(Ymax)
                    .CrossesAt = .MinimumScale
                End If
            End With
        Next
        
        '�}��̏����ݒ�
        .HasLegend = True
        With .Legend
            .IncludeInLayout = False
            .Format.Fill.ForeColor.RGB = rgbWhite
            .Format.Line.ForeColor.RGB = rgbBlack
            With .Format.Shadow
                .Type = msoShadow21
                .Size = 102
            End With
        End With
        
        '�n��̏����ݒ�
'        Dim j As Long, sc As Object
'        j = 0
'        For Each sc In .SeriesCollection
'            sc.MarkerStyle = Array(8, 3, 1, 2)(j Mod 4)
'            If UBound(sc.Values) >= 10 ^ 2 Then sc.MarkerStyle = xlNone
'            sc.Format.Line.DashStyle = Array(1, 4, 5, 6, 7, 8, 2)(j Mod 7)
'            sc.MarkerSize = Marker_Size
'            sc.Format.Line.ForeColor.RGB = rgbBlack
'            sc.Format.Fill.ForeColor.RGB = rgbBlack
'            j = j + 1
'        Next
        
        cht.ShapeRange.Line.Visible = msoFalse
        With .PlotArea.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = rgbBlack
        End With
    End With
    
     ' �o��
    Set Plot = cht
   
End Function

Private Function MyMin(ParamArray arr())
    Dim arr1, arr2
    arr1 = arr(0)
    arr2 = arr(1)
    MyMin = WorksheetFunction.Min(arr1, arr2)
End Function
Private Function MyMax(ParamArray arr())
    Dim arr1, arr2
    arr1 = arr(0)
    arr2 = arr(1)
    MyMax = WorksheetFunction.Max(arr1, arr2)
End Function
