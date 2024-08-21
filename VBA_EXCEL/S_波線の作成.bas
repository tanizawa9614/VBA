Attribute VB_Name = "S_îgê¸ÇÃçÏê¨"
Option Explicit

Sub îgê¸ÇÃçÏê¨()

    Dim startX As Double, startY As Double
    Dim endX As Double, endY As Double
    Dim n As Long, i As Long
    Dim H As Double
    Dim x As Double, y As Double
    Dim a As Double, LW As Double
    
    startX = 10
    endX = 100
    startY = 10
    endY = 10
    n = 60 'îgÇÃêî
    H = 1 'îgçÇ
    LW = 0.1 'ê¸ÇÃïù
    
    a = (endY - startY) / (endX - startX)
    
    With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, startX, startY)
        For i = 1 To 4 * n
            x = Abs(startX + (endX - startX) / (4 * n) * i)
            y = Abs(a * x + startY + sighn(i) * H)
            .AddNodes msoSegmentCurve, msoEditingAuto, x, y
        Next
        .ConvertToShape
    End With
    With ActiveSheet.Shapes
        With .Range(.Count).Line
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = LW
        End With
    End With
End Sub

Private Function sighn(i As Long) As Long
    Dim ans As Long
    Dim tmp As Long
    tmp = modi(i, 4)
    Select Case tmp
    Case 1
        ans = 1
    Case 2, 0
        ans = 0
    Case 3
        ans = -1
    End Select
    sighn = ans
End Function
Private Function modi(i As Long, j As Long) As Long
    Dim tmp As Long
    tmp = Int(i / j)
    modi = i - tmp * j
End Function
