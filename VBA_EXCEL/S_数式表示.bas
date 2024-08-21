Attribute VB_Name = "S_数式表示"
Option Explicit
Sub 数式削除()
   Dim i As Long
   On Error Resume Next
   For i = 1 To 3
      ActiveSheet.Shapes(1).Delete
   Next
End Sub

Sub 数式表示()
Attribute 数式表示.VB_ProcData.VB_Invoke_Func = " \n14"
   Application.ScreenUpdating = False
   Dim R As Range, Fun, i As Long, j As Long
   Dim cnt As Long
   Dim sta As Long, buf, FSize As Long
   Dim addText As Object, addTri As Object
   Dim addTri2 As Object, Gname As Object
   Dim addCon As Object
   Dim CSt(12), Ld As Long, Td As Long
   CSt(0) = ",": CSt(1) = "(": CSt(2) = ")"
   CSt(3) = "+": CSt(4) = "-": CSt(5) = "*"
   CSt(6) = "/": CSt(7) = "=": CSt(8) = ">"
   CSt(9) = "<": CSt(10) = "?": CSt(11) = """"
   CSt(12) = "&"
   
   For Each R In Selection
   '   TextBoxの追加
      With R
         ActiveSheet.Shapes.AddLabel _
         (msoTextOrientationHorizontal, _
         R.Left, R.Top, 1, 1).Select
      End With
      Set addText = Selection
      Selection.OnAction = "ダミー"
   
   '   Fontの設定
      With WorksheetFunction
         FSize = _
         .Max(.Min(-0.17 * ActiveWindow.Zoom + 47, 46), 16)
         Td = (cnt + 1) * R.ColumnWidth * FSize * 0.18
         Ld = (cnt + 1) * R.RowHeight * FSize * 0.18
      End With
      With addText.ShapeRange(1).TextFrame2
         With .TextRange
            .Text = R.Formula
            .ParagraphFormat.Alignment = msoAlignCenter
            With .Font
               .Name = "ＭＳ ゴシック"
               .NameFarEast = "ＭＳ ゴシック"
               .Fill.ForeColor.RGB = RGB(0, 150, 80)
               .Size = FSize
            End With
         End With
         .WordWrap = msoFalse
      End With
      
   '  塗りつぶしと枠線
      With addText.ShapeRange(1)
         With .Fill
         .ForeColor.RGB = RGB(255, 255, 220)
         End With
         With .Line
            .Visible = msoTrue
            .ForeColor.RGB = 0
            .Weight = 0.75
         End With
      End With
      
   '   関数に色付け
      Fun = D_function(R.Formula)
      If Not IsEmpty(Fun) Then
         For i = LBound(Fun) To UBound(Fun)
            buf = Split(R.Formula, Fun(i))
            sta = 0
            For j = 0 To UBound(buf) - 1
               sta = sta + Len(buf(j))
               addText.ShapeRange(1).TextFrame2. _
                  TextRange.Characters(sta + 1, Len(Fun(i))) _
                  .Font.Fill.ForeColor.RGB = RGB(0, 0, 255)
               sta = sta + Len(Fun(i))
            Next j
         Next
      End If
   
   '   特定文字に色付け
      For i = 0 To UBound(CSt)
         buf = Split(R.Formula, CSt(i))
         sta = 0
            For j = 0 To UBound(buf) - 1
               sta = sta + Len(buf(j))
               With addText.ShapeRange(1).TextFrame2. _
                  TextRange.Characters(sta + 1, Len(CSt(i))) _
                  .Font.Fill.ForeColor
                  If i >= 0 And i <= 2 Then
                     .RGB = RGB(255, 0, 0)
                  Else
                     .RGB = 0
                  End If
               End With
               sta = sta + Len(CSt(i))
            Next j
      Next i
      
   '   三角形を作成
      ActiveSheet.Shapes.AddShape(msoShapeIsoscelesTriangle, R.Offset(, 1).Left - 3, _
         R.Top, 3, 3).Select
      Set addTri = Selection
      With addTri.ShapeRange
         .Adjustments.Item(1) = 1
         .Flip msoFlipVertical
         .Fill.ForeColor.RGB = RGB(51, 51, 255)
         .Line.Visible = msoFalse
         .Duplicate.Select
      End With
      Set addTri2 = Selection
      ActiveSheet.Shapes.Range(Array(addTri2.Name, addText.Name)).Select
      With Selection.ShapeRange
         .Align msoAlignRights, msoFalse
         .Align msoAlignTops, msoFalse
         Set Gname = .Group
         Gname.ZOrder msoSendToBack
         Gname.IncrementLeft Ld
         Gname.IncrementTop Td
      End With
      
   '   直線を追加
      ActiveSheet.Shapes.AddConnector(msoConnectorStraight, R.Offset(, 1).Left, _
         R.Top, 120.1941732283, 82.7184251969).Select
      Set addCon = Selection
      With addCon.ShapeRange.Line
         .EndArrowheadStyle = msoArrowheadTriangle
         .EndArrowheadLength = msoArrowheadShort
         .EndArrowheadWidth = msoArrowheadNarrow
         .Weight = 0.6
         .ForeColor.RGB = 0
      End With
      With addCon.ShapeRange.ConnectorFormat
         .BeginConnect ActiveSheet.Shapes( _
           addTri2.Name), 4
         .EndConnect ActiveSheet.Shapes( _
           addTri.Name), 4
      End With
      With ActiveSheet
         .Shapes(addTri.Name).ZOrder msoSendToBack
         .Shapes(addTri2.Name).ZOrder msoSendToBack
         .Shapes(addCon.Name).ZOrder msoSendToBack
      End With
      cnt = cnt + 1
   Next R
   ActiveSheet.Shapes.Range(Array(Gname.Name)).Select
   Application.ScreenUpdating = True
End Sub

Private Function D_function(F1 As String) As Variant
   Dim n As Long, A, buf As String, i As Long
   Dim F As String
   Dim j As Long
   Dim CSt(7)
   CSt(0) = "("
   CSt(1) = ")"
   CSt(2) = ","
   CSt(3) = "+"
   CSt(4) = "-"
   CSt(5) = "*"
   CSt(6) = "/"
   CSt(7) = "="
   
   F = F1
   
   If InStr(F, "=") = 0 Or InStr(F, "(") = 0 _
   Then Exit Function
   
   For i = LBound(CSt) To UBound(CSt)
      F = Replace(F, CSt(i), vbTab)
   Next
   A = Split(F, vbTab)
   For i = LBound(A) To UBound(A)
      buf = A(i)
      If buf <> "" And IsNumeric(buf) = False _
         And Mid(F1, InStr(F1, buf) + Len(buf), 1) = "(" Then
         If アドレスかどうか(buf) = False Then
            A(j) = buf
            j = j + 1
         End If
      End If
   Next
   If i >= 1 Then
      ReDim Preserve A(j - 1)
   Else
      A = ""
   End If
   D_function = A
End Function

Private Function アドレスかどうか(S As String) As Boolean
   On Error Resume Next
   Dim buf
   buf = Range(S).Address(False, False)
   If buf <> "" Or S = "TRUE" Or S = "FALSE" Then
      アドレスかどうか = True
      Exit Function
   End If
   アドレスかどうか = False
End Function

Private Sub ダミー()
   
End Sub
