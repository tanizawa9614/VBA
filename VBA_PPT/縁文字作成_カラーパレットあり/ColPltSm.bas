Attribute VB_Name = "ColPltSm"
' •W€ƒ‚ƒWƒ…[ƒ‹
Option Explicit
Public ColVal(11, 11) As Long
Public ColName(11, 11) As String
Public SizeMe As Double
Public myImage0 As String
Public AnsCol As Long

Public Sub CallColPalette()
    ColPltUF.Show
End Sub

Public Function IsInObj(Obj As MSForms.CommandButton, ByVal X As Single, ByVal Y As Single) As Boolean
    Dim x0 As Double
    Dim y0 As Double
    Dim x1 As Double
    Dim y1 As Double
    
    With Obj
        x0 = .Left
        x1 = .Left + .Width
        y0 = .Top
        y1 = .Top + .Height
    End With
        
    IsInObj = False
    If X >= x0 And X <= x1 And Y >= y0 And Y <= y1 Then IsInObj = True
    
End Function

Public Sub ColorBlue(Obj As MSForms.CommandButton)
    With Obj
        .BackColor = RGB(201, 236, 255)
    End With
End Sub

Public Sub ColorGray(Obj As MSForms.CommandButton)
    With Obj
        .BackColor = rgbWhiteSmoke
    End With
End Sub


Sub MakeColorArray()
    ColVal(0, 0) = rgbWhite
    ColVal(0, 1) = rgbAzure
    ColVal(0, 2) = rgbAqua
    ColVal(0, 3) = rgbGhostWhite
    ColVal(0, 4) = rgbAliceBlue
    ColVal(0, 5) = rgbDeepSkyBlue
    ColVal(0, 6) = rgbDodgerBlue
    ColVal(0, 7) = rgbFuchsia
    ColVal(0, 8) = rgbBlue
    ColVal(0, 9) = rgbMintCream
    ColVal(0, 10) = rgbSnow
    ColVal(0, 11) = rgbLavender
    ColVal(1, 0) = rgbLightSkyBlue
    ColVal(1, 1) = rgbWhiteSmoke
    ColVal(1, 2) = rgbLavenderBlush
    ColVal(1, 3) = rgbIvory
    ColVal(1, 4) = rgbHoneydew
    ColVal(1, 5) = rgbFloralWhite
    ColVal(1, 6) = rgbSeashell
    ColVal(1, 7) = rgbPaleTurquoise
    ColVal(1, 8) = rgbViolet
    ColVal(1, 9) = rgbMediumSlateBlue
    ColVal(1, 10) = rgbCornflowerBlue
    ColVal(1, 11) = rgbSkyBlue
    ColVal(2, 0) = rgbOldLace
    ColVal(2, 1) = rgbLinen
    ColVal(2, 2) = rgbPowderBlue
    ColVal(2, 3) = rgbLightBlue
    ColVal(2, 4) = rgbBlueViolet
    ColVal(2, 5) = rgbMistyRose
    ColVal(2, 6) = rgbRoyalBlue
    ColVal(2, 7) = rgbLightYellow
    ColVal(2, 8) = rgbLightSteelBlue
    ColVal(2, 9) = rgbPlum
    ColVal(2, 10) = rgbCornsilk
    ColVal(2, 11) = rgbBeige
    ColVal(3, 0) = rgbGainsboro
    ColVal(3, 1) = rgbMediumPurple
    ColVal(3, 2) = rgbThistle
    ColVal(3, 3) = rgbAntiqueWhite
    ColVal(3, 4) = rgbOrchid
    ColVal(3, 5) = rgbPapayaWhip
    ColVal(3, 6) = rgbAquamarine
    ColVal(3, 7) = rgbLightGray
    ColVal(3, 8) = rgbLightGray
    ColVal(3, 9) = rgbMediumOrchid
    ColVal(3, 10) = rgbDarkViolet
    ColVal(3, 11) = rgbLightGoldenrodYellow
    ColVal(4, 0) = rgbDarkTurquoise
    ColVal(4, 1) = rgbTurquoise
    ColVal(4, 2) = rgbLemonChiffon
    ColVal(4, 3) = rgbBlanchedAlmond
    ColVal(4, 4) = rgbSlateBlue
    ColVal(4, 5) = rgbMediumBlue
    ColVal(4, 6) = rgbMediumTurquoise
    ColVal(4, 7) = rgbDarkOrchid
    ColVal(4, 8) = rgbPink
    ColVal(4, 9) = rgbBisque
    ColVal(4, 10) = rgbLightPink
    ColVal(4, 11) = rgbSilver
    ColVal(5, 0) = rgbPeachPuff
    ColVal(5, 1) = rgbMoccasin
    ColVal(5, 2) = rgbSteelBlue
    ColVal(5, 3) = rgbHotPink
    ColVal(5, 4) = rgbWheat
    ColVal(5, 5) = rgbNavajoWhite
    ColVal(5, 6) = rgbMediumAquamarine
    ColVal(5, 7) = rgbLightSeaGreen
    ColVal(5, 8) = rgbDarkGray
    ColVal(5, 9) = rgbDarkGray
    ColVal(5, 10) = rgbCadetBlue
    ColVal(5, 11) = rgbMediumSpringGreen
    ColVal(6, 0) = rgbLightSlateGray
    ColVal(6, 1) = rgbPaleGreen
    ColVal(6, 2) = rgbPaleVioletRed
    ColVal(6, 3) = rgbDeepPink
    ColVal(6, 4) = rgbLightGreen
    ColVal(6, 5) = rgbSlateGray
    ColVal(6, 6) = rgbDarkSeaGreen
    ColVal(6, 7) = rgbRosyBrown
    ColVal(6, 8) = rgbKhaki
    ColVal(6, 9) = rgbTan
    ColVal(6, 10) = rgbDarkCyan
    ColVal(6, 11) = rgbDarkCyan
    ColVal(7, 0) = rgbDarkSlateBlue
    ColVal(7, 1) = rgbDarkMagenta
    ColVal(7, 2) = rgbDarkBlue
    ColVal(7, 3) = rgbBurlyWood
    ColVal(7, 4) = rgbMediumVioletRed
    ColVal(7, 5) = rgbIndigo
    ColVal(7, 6) = rgbLightCoral
    ColVal(7, 7) = rgbGray
    ColVal(7, 8) = rgbGray
    ColVal(7, 9) = rgbTeal
    ColVal(7, 10) = rgbPurple
    ColVal(7, 11) = rgbNavy
    ColVal(8, 0) = rgbNavy
    ColVal(8, 1) = rgbSpringGreen
    ColVal(8, 2) = rgbLightSalmon
    ColVal(8, 3) = rgbDarkSalmon
    ColVal(8, 4) = rgbSalmon
    ColVal(8, 5) = rgbMediumSeaGreen
    ColVal(8, 6) = rgbMidnightBlue
    ColVal(8, 7) = rgbPaleGoldenrod
    ColVal(8, 8) = rgbDarkKhaki
    ColVal(8, 9) = rgbDimGray
    ColVal(8, 10) = rgbDimGray
    ColVal(8, 11) = rgbSandyBrown
    ColVal(9, 0) = rgbIndianRed
    ColVal(9, 1) = rgbSeaGreen
    ColVal(9, 2) = rgbCoral
    ColVal(9, 3) = rgbDarkSlateGray
    ColVal(9, 4) = rgbDarkSlateGray
    ColVal(9, 5) = rgbTomato
    ColVal(9, 6) = rgbPeru
    ColVal(9, 7) = rgbCrimson
    ColVal(9, 8) = rgbYellowGreen
    ColVal(9, 9) = rgbLimeGreen
    ColVal(9, 10) = rgbGreenYellow
    ColVal(9, 11) = rgbDarkOliveGreen
    ColVal(10, 0) = rgbSienna
    ColVal(10, 1) = rgbBrown
    ColVal(10, 2) = rgbOliveDrab
    ColVal(10, 3) = rgbForestGreen
    ColVal(10, 4) = rgbFireBrick
    ColVal(10, 5) = rgbGoldenrod
    ColVal(10, 6) = rgbDarkGoldenrod
    ColVal(10, 7) = rgbYellow
    ColVal(10, 8) = rgbChartreuse
    ColVal(10, 9) = rgbLime
    ColVal(10, 10) = rgbLawnGreen
    ColVal(10, 11) = rgbGold
    ColVal(11, 0) = rgbOrange
    ColVal(11, 1) = rgbDarkOrange
    ColVal(11, 2) = rgbOlive
    ColVal(11, 3) = rgbGreen
    ColVal(11, 4) = rgbDarkGreen
    ColVal(11, 5) = rgbOrangeRed
    ColVal(11, 6) = rgbRed
    ColVal(11, 7) = rgbDarkRed
    ColVal(11, 8) = rgbMaroon
    ColVal(11, 9) = rgbBlack
    ColVal(11, 10) = rgbBlack
    ColVal(11, 11) = rgbBlack
    ColName(0, 0) = "rgbWhite"
    ColName(0, 1) = "rgbAzure"
    ColName(0, 2) = "rgbAqua"
    ColName(0, 3) = "rgbGhostWhite"
    ColName(0, 4) = "rgbAliceBlue"
    ColName(0, 5) = "rgbDeepSkyBlue"
    ColName(0, 6) = "rgbDodgerBlue"
    ColName(0, 7) = "rgbFuchsia"
    ColName(0, 8) = "rgbBlue"
    ColName(0, 9) = "rgbMintCream"
    ColName(0, 10) = "rgbSnow"
    ColName(0, 11) = "rgbLavender"
    ColName(1, 0) = "rgbLightSkyBlue"
    ColName(1, 1) = "rgbWhiteSmoke"
    ColName(1, 2) = "rgbLavenderBlush"
    ColName(1, 3) = "rgbIvory"
    ColName(1, 4) = "rgbHoneydew"
    ColName(1, 5) = "rgbFloralWhite"
    ColName(1, 6) = "rgbSeashell"
    ColName(1, 7) = "rgbPaleTurquoise"
    ColName(1, 8) = "rgbViolet"
    ColName(1, 9) = "rgbMediumSlateBlue"
    ColName(1, 10) = "rgbCornflowerBlue"
    ColName(1, 11) = "rgbSkyBlue"
    ColName(2, 0) = "rgbOldLace"
    ColName(2, 1) = "rgbLinen"
    ColName(2, 2) = "rgbPowderBlue"
    ColName(2, 3) = "rgbLightBlue"
    ColName(2, 4) = "rgbBlueViolet"
    ColName(2, 5) = "rgbMistyRose"
    ColName(2, 6) = "rgbRoyalBlue"
    ColName(2, 7) = "rgbLightYellow"
    ColName(2, 8) = "rgbLightSteelBlue"
    ColName(2, 9) = "rgbPlum"
    ColName(2, 10) = "rgbCornsilk"
    ColName(2, 11) = "rgbBeige"
    ColName(3, 0) = "rgbGainsboro"
    ColName(3, 1) = "rgbMediumPurple"
    ColName(3, 2) = "rgbThistle"
    ColName(3, 3) = "rgbAntiqueWhite"
    ColName(3, 4) = "rgbOrchid"
    ColName(3, 5) = "rgbPapayaWhip"
    ColName(3, 6) = "rgbAquamarine"
    ColName(3, 7) = "rgbLightGray"
    ColName(3, 8) = "rgbLightGray"
    ColName(3, 9) = "rgbMediumOrchid"
    ColName(3, 10) = "rgbDarkViolet"
    ColName(3, 11) = "rgbLightGoldenrodYellow"
    ColName(4, 0) = "rgbDarkTurquoise"
    ColName(4, 1) = "rgbTurquoise"
    ColName(4, 2) = "rgbLemonChiffon"
    ColName(4, 3) = "rgbBlanchedAlmond"
    ColName(4, 4) = "rgbSlateBlue"
    ColName(4, 5) = "rgbMediumBlue"
    ColName(4, 6) = "rgbMediumTurquoise"
    ColName(4, 7) = "rgbDarkOrchid"
    ColName(4, 8) = "rgbPink"
    ColName(4, 9) = "rgbBisque"
    ColName(4, 10) = "rgbLightPink"
    ColName(4, 11) = "rgbSilver"
    ColName(5, 0) = "rgbPeachPuff"
    ColName(5, 1) = "rgbMoccasin"
    ColName(5, 2) = "rgbSteelBlue"
    ColName(5, 3) = "rgbHotPink"
    ColName(5, 4) = "rgbWheat"
    ColName(5, 5) = "rgbNavajoWhite"
    ColName(5, 6) = "rgbMediumAquamarine"
    ColName(5, 7) = "rgbLightSeaGreen"
    ColName(5, 8) = "rgbDarkGray"
    ColName(5, 9) = "rgbDarkGray"
    ColName(5, 10) = "rgbCadetBlue"
    ColName(5, 11) = "rgbMediumSpringGreen"
    ColName(6, 0) = "rgbLightSlateGray"
    ColName(6, 1) = "rgbPaleGreen"
    ColName(6, 2) = "rgbPaleVioletRed"
    ColName(6, 3) = "rgbDeepPink"
    ColName(6, 4) = "rgbLightGreen"
    ColName(6, 5) = "rgbSlateGray"
    ColName(6, 6) = "rgbDarkSeaGreen"
    ColName(6, 7) = "rgbRosyBrown"
    ColName(6, 8) = "rgbKhaki"
    ColName(6, 9) = "rgbTan"
    ColName(6, 10) = "rgbDarkCyan"
    ColName(6, 11) = "rgbDarkCyan"
    ColName(7, 0) = "rgbDarkSlateBlue"
    ColName(7, 1) = "rgbDarkMagenta"
    ColName(7, 2) = "rgbDarkBlue"
    ColName(7, 3) = "rgbBurlyWood"
    ColName(7, 4) = "rgbMediumVioletRed"
    ColName(7, 5) = "rgbIndigo"
    ColName(7, 6) = "rgbLightCoral"
    ColName(7, 7) = "rgbGray"
    ColName(7, 8) = "rgbGray"
    ColName(7, 9) = "rgbTeal"
    ColName(7, 10) = "rgbPurple"
    ColName(7, 11) = "rgbNavy"
    ColName(8, 0) = "rgbNavy"
    ColName(8, 1) = "rgbSpringGreen"
    ColName(8, 2) = "rgbLightSalmon"
    ColName(8, 3) = "rgbDarkSalmon"
    ColName(8, 4) = "rgbSalmon"
    ColName(8, 5) = "rgbMediumSeaGreen"
    ColName(8, 6) = "rgbMidnightBlue"
    ColName(8, 7) = "rgbPaleGoldenrod"
    ColName(8, 8) = "rgbDarkKhaki"
    ColName(8, 9) = "rgbDimGray"
    ColName(8, 10) = "rgbDimGray"
    ColName(8, 11) = "rgbSandyBrown"
    ColName(9, 0) = "rgbIndianRed"
    ColName(9, 1) = "rgbSeaGreen"
    ColName(9, 2) = "rgbCoral"
    ColName(9, 3) = "rgbDarkSlateGray"
    ColName(9, 4) = "rgbDarkSlateGray"
    ColName(9, 5) = "rgbTomato"
    ColName(9, 6) = "rgbPeru"
    ColName(9, 7) = "rgbCrimson"
    ColName(9, 8) = "rgbYellowGreen"
    ColName(9, 9) = "rgbLimeGreen"
    ColName(9, 10) = "rgbGreenYellow"
    ColName(9, 11) = "rgbDarkOliveGreen"
    ColName(10, 0) = "rgbSienna"
    ColName(10, 1) = "rgbBrown"
    ColName(10, 2) = "rgbOliveDrab"
    ColName(10, 3) = "rgbForestGreen"
    ColName(10, 4) = "rgbFireBrick"
    ColName(10, 5) = "rgbGoldenrod"
    ColName(10, 6) = "rgbDarkGoldenrod"
    ColName(10, 7) = "rgbYellow"
    ColName(10, 8) = "rgbChartreuse"
    ColName(10, 9) = "rgbLime"
    ColName(10, 10) = "rgbLawnGreen"
    ColName(10, 11) = "rgbGold"
    ColName(11, 0) = "rgbOrange"
    ColName(11, 1) = "rgbDarkOrange"
    ColName(11, 2) = "rgbOlive"
    ColName(11, 3) = "rgbGreen"
    ColName(11, 4) = "rgbDarkGreen"
    ColName(11, 5) = "rgbOrangeRed"
    ColName(11, 6) = "rgbRed"
    ColName(11, 7) = "rgbDarkRed"
    ColName(11, 8) = "rgbMaroon"
    ColName(11, 9) = "rgbBlack"
    ColName(11, 10) = "rgbBlack"
    ColName(11, 11) = "rgbBlack"
End Sub


Private Sub InteriorColorChange()
    Dim rg As Range
    Dim i As Long
    Dim j As Long
    
    Set rg = Selection
    For i = 1 To rg.Rows.Count
        For j = 1 To rg.Columns.Count
            With rg.Resize(1, 1).Offset(i - 1, j - 1)
                .Interior.Color = .Value
            End With
        Next
    Next
End Sub


