Attribute VB_Name = "S_�ʒu���킹0"
Option Explicit
Dim shps()
Dim Groupflg As Boolean
Dim GroupArray() As String
Dim GroupName As String
Dim SelectedGroupArray() As String

Sub ������()
    Dim L As Double, i As Long
    Call �O���[�v���̈ꎞ����
'    On Error Resume Next
    With ActiveWindow.Selection
        L = .ShapeRange(.ShapeRange.Count).Left
        For i = 1 To .ShapeRange.Count - 1
            .ShapeRange(i).Left = L
        Next
        If .ShapeRange.Count = 1 Then
            .ShapeRange(1).Left = 0 'ActivePresentation.PageSetup
        End If
    End With
    Call �O���[�v���̕���
End Sub

Sub ���E��������()
    Dim L As Double, W As Double
    Dim M As Long, i As Long
    
    Call �O���[�v���̈ꎞ����
'    On Error Resume Next
    With ActiveWindow.Selection
        L = .ShapeRange(.ShapeRange.Count).Left
        W = .ShapeRange(.ShapeRange.Count).Width
        M = L + W * 0.5
        For i = 1 To .ShapeRange.Count - 1
            .ShapeRange(i).Left = M - .ShapeRange(i).Width * 0.5
        Next
        If .ShapeRange.Count = 1 Then
            .ShapeRange(1).Left = ActivePresentation.PageSetup.SlideWidth * 0.5 - W * 0.5
        End If
    End With
    Call �O���[�v���̕���
End Sub

Sub �E����()
   Dim L As Double, W As Double
   Dim M As Long, i As Long
   
   Call �O���[�v���̈ꎞ����
'   On Error Resume Next
   With ActiveWindow.Selection
      L = .ShapeRange(.ShapeRange.Count).Left
      W = .ShapeRange(.ShapeRange.Count).Width
      M = L + W
      For i = 1 To .ShapeRange.Count - 1
         .ShapeRange(i).Left = M - .ShapeRange(i).Width
      Next
      If .ShapeRange.Count = 1 Then
            .ShapeRange(1).Left = ActivePresentation.PageSetup.SlideWidth - W
      End If
   End With
   Call �O���[�v���̕���
End Sub

Sub �㑵��()
   Dim T As Double, i As Long
   
   Call �O���[�v���̈ꎞ����
'   On Error Resume Next
   With ActiveWindow.Selection
      T = .ShapeRange(.ShapeRange.Count).Top
      For i = 1 To .ShapeRange.Count - 1
         .ShapeRange(i).Top = T
      Next
      If .ShapeRange.Count = 1 Then
            .ShapeRange(1).Top = 0
      End If
   End With
   Call �O���[�v���̕���
End Sub

Sub �㉺��������()
   Dim T As Double, H As Double
   Dim M As Long, i As Long
   
   Call �O���[�v���̈ꎞ����
'   On Error Resume Next
   With ActiveWindow.Selection
      T = .ShapeRange(.ShapeRange.Count).Top
      H = .ShapeRange(.ShapeRange.Count).Height
      M = T + H * 0.5
      For i = 1 To .ShapeRange.Count - 1
         .ShapeRange(i).Top = M - .ShapeRange(i).Height * 0.5
      Next
      If .ShapeRange.Count = 1 Then
            .ShapeRange(1).Top = ActivePresentation.PageSetup.SlideHeight * 0.5 - H * 0.5
      End If
   End With
   Call �O���[�v���̕���
End Sub

Sub ������()
   Dim T As Double, H As Double
   Dim M As Long, i As Long
   
   Call �O���[�v���̈ꎞ����
'   On Error Resume Next
   With ActiveWindow.Selection
      T = .ShapeRange(.ShapeRange.Count).Top
      H = .ShapeRange(.ShapeRange.Count).Height
      M = T + H
      For i = 1 To .ShapeRange.Count - 1
         .ShapeRange(i).Top = M - .ShapeRange(i).Height
      Next
      If .ShapeRange.Count = 1 Then
            .ShapeRange(1).Top = ActivePresentation.PageSetup.SlideHeight - H
      End If
   End With
   Call �O���[�v���̕���
End Sub

Private Sub �O���[�v���̈ꎞ����()
    Dim shp As Shape
    Dim gshp As Shape
    Dim n As Long
    Dim ns As Long
    Dim i As Long
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    
    ' �O���[�v�����̐}�`�Ȃ�ꎞ�I�ɃO���[�v������������
    If shp.Type = msoGroup And ActiveWindow.Selection.ShapeRange.Count = 1 Then
        Groupflg = True
        n = shp.GroupItems.Count
        ReDim GroupArray(1 To n)  '�O���[�v������Ă��邷�ׂĂ̐}�`���擾
        For i = 1 To n
            GroupArray(i) = shp.GroupItems(i).Name
        Next
        GroupName = shp.Name  '�O���[�v������Ă���}�`�̖��O
        shp.Ungroup  '�O���[�v������
        On Error Resume Next
        ns = ActiveWindow.Selection.ShapeRange.Count
        On Error GoTo 0
        If ns = 0 Then
            Groupflg = True
            Call �O���[�v���̕���
            Groupflg = False
            Exit Sub
        End If
        ReDim SelectedGroupArray(1 To ns)  '�O���[�v�����őI������Ă����}�`
        For i = 1 To ns
            SelectedGroupArray(i) = ActiveWindow.Selection.ShapeRange(i).Name
        Next
    Else
        Groupflg = False
    End If
End Sub

Private Sub �O���[�v���̕���()
    Dim i As Long
    If Groupflg = False Then Exit Sub  '���X���O���[�v������Ă����ꍇ�̂ݏ���
    
    Dim Si As Long, Sld As Slide
    Si = ActiveWindow.Selection.SlideRange.SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
     
    ' �}�`�𕡐��I��
    For i = 1 To UBound(GroupArray)
        If i = 1 Then
            Sld.Shapes(GroupArray(i)).Select '�}�`��I��
        Else
            Sld.Shapes(GroupArray(i)).Select Replace:=False '�}�`���u�ǉ�]
        End If
    Next
    
    '�@�O���[�v������
    With ActiveWindow.Selection.ShapeRange.Group
        .Name = GroupName
        .Select
    End With
    
    Dim n As Long
    On Error Resume Next
    n = UBound(SelectedGroupArray)
    On Error GoTo 0
    
    '���X�I������Ă����}�`��I�����Ȃ���
    For i = 1 To n
        If i = 1 Then
            Sld.Shapes(GroupName).GroupItems(SelectedGroupArray(i)).Select '�}�`��I��
        Else
            Sld.Shapes(GroupName).GroupItems(SelectedGroupArray(i)).Select Replace:=False '�}�`���u�ǉ�] '�}�`��I��
        End If
    Next
    
End Sub
