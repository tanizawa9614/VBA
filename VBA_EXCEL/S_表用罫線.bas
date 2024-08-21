Attribute VB_Name = "S_表用罫線"
Option Explicit

Sub 表用罫線()
   Dim Target As Range
   Dim buf As Range
   Dim i As Long
   Dim R As Long, C As Long
   On Error GoTo L1
   Call AppActivate(ThisWorkbook.Name)
   Set Target = Application.InputBox("対象セルを選択してください", Type:=8)
   Application.ScreenUpdating = False
   
   R = Target.Rows.Count
   C = Target.Columns.Count
   Target.Borders.LineStyle = xlNone
'上端
   With Target.Resize(1)
      With .Borders(xlEdgeBottom) '上端下側
'         .LineStyle = xlContinuous
         .LineStyle = xlDouble
         .Weight = xlThick
      End With
      With .Borders(xlEdgeTop) '最上端
         .LineStyle = xlContinuous
'         .LineStyle = xlDouble
         .Weight = xlThick
      End With
'下端
   End With
   With Target.Resize(1).Offset(R - 1)
      With .Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThick
      End With
   End With
'縦線
   For i = 1 To C - 1
      With Target.Resize(, 1).Offset(, i).Borders(xlEdgeLeft)
         .LineStyle = xlContinuous
      End With
   Next
'横線
   For i = 2 To R - 1
      With Target.Resize(1).Offset(i).Borders(xlEdgeTop)
         .LineStyle = xlDot
      End With
   Next
L1:
   Application.ScreenUpdating = True
End Sub
