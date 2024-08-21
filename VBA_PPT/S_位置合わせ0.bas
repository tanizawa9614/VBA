Attribute VB_Name = "S_位置合わせ0"
Option Explicit
Dim shps()
Dim Groupflg As Boolean
Dim GroupArray() As String
Dim GroupName As String
Dim SelectedGroupArray() As String

Sub 左揃え()
    Dim L As Double, i As Long
    Call グループ化の一時解除
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
    Call グループ化の復元
End Sub

Sub 左右中央揃え()
    Dim L As Double, W As Double
    Dim M As Long, i As Long
    
    Call グループ化の一時解除
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
    Call グループ化の復元
End Sub

Sub 右揃え()
   Dim L As Double, W As Double
   Dim M As Long, i As Long
   
   Call グループ化の一時解除
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
   Call グループ化の復元
End Sub

Sub 上揃え()
   Dim T As Double, i As Long
   
   Call グループ化の一時解除
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
   Call グループ化の復元
End Sub

Sub 上下中央揃え()
   Dim T As Double, H As Double
   Dim M As Long, i As Long
   
   Call グループ化の一時解除
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
   Call グループ化の復元
End Sub

Sub 下揃え()
   Dim T As Double, H As Double
   Dim M As Long, i As Long
   
   Call グループ化の一時解除
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
   Call グループ化の復元
End Sub

Private Sub グループ化の一時解除()
    Dim shp As Shape
    Dim gshp As Shape
    Dim n As Long
    Dim ns As Long
    Dim i As Long
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    
    ' グループ化内の図形なら一時的にグループ化を解除する
    If shp.Type = msoGroup And ActiveWindow.Selection.ShapeRange.Count = 1 Then
        Groupflg = True
        n = shp.GroupItems.Count
        ReDim GroupArray(1 To n)  'グループ化されているすべての図形を取得
        For i = 1 To n
            GroupArray(i) = shp.GroupItems(i).Name
        Next
        GroupName = shp.Name  'グループ化されている図形の名前
        shp.Ungroup  'グループ化解除
        On Error Resume Next
        ns = ActiveWindow.Selection.ShapeRange.Count
        On Error GoTo 0
        If ns = 0 Then
            Groupflg = True
            Call グループ化の復元
            Groupflg = False
            Exit Sub
        End If
        ReDim SelectedGroupArray(1 To ns)  'グループ化内で選択されていた図形
        For i = 1 To ns
            SelectedGroupArray(i) = ActiveWindow.Selection.ShapeRange(i).Name
        Next
    Else
        Groupflg = False
    End If
End Sub

Private Sub グループ化の復元()
    Dim i As Long
    If Groupflg = False Then Exit Sub  '元々がグループ化されていた場合のみ処理
    
    Dim Si As Long, Sld As Slide
    Si = ActiveWindow.Selection.SlideRange.SlideIndex
    Set Sld = ActivePresentation.Slides(Si)
     
    ' 図形を複数選択
    For i = 1 To UBound(GroupArray)
        If i = 1 Then
            Sld.Shapes(GroupArray(i)).Select '図形を選択
        Else
            Sld.Shapes(GroupArray(i)).Select Replace:=False '図形を「追加]
        End If
    Next
    
    '　グループ化復元
    With ActiveWindow.Selection.ShapeRange.Group
        .Name = GroupName
        .Select
    End With
    
    Dim n As Long
    On Error Resume Next
    n = UBound(SelectedGroupArray)
    On Error GoTo 0
    
    '元々選択されていた図形を選択しなおす
    For i = 1 To n
        If i = 1 Then
            Sld.Shapes(GroupName).GroupItems(SelectedGroupArray(i)).Select '図形を選択
        Else
            Sld.Shapes(GroupName).GroupItems(SelectedGroupArray(i)).Select Replace:=False '図形を「追加] '図形を選択
        End If
    Next
    
End Sub
