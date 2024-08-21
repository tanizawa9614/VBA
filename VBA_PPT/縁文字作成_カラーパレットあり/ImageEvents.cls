VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImageEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class
Option Explicit
Private WithEvents myImage As MSForms.Image
Attribute myImage.VB_VarHelpID = -1


Public Sub SetImage(NewImage As MSForms.Image)
    Set myImage = NewImage
End Sub

Private Sub myImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim ijstr As String
    Dim i As Long
    Dim j As Long
    
    BorderInitialize
    
    With myImage
        .BorderColor = RGB(255, 255, 255)
        .BorderStyle = 1
        ijstr = Replace(.Name, "Image", "")
    End With
    
    i = Val(Left(ijstr, 2))
    j = Val(Right(ijstr, 2))
    
    With ColPltUF.Controls("LabalCurrent")
        .Caption = ColName(i, j)
        .Width = 300
        .AutoSize = True
    End With
    
    With ColPltUF.Controls("ImageCurrent")
        .BackColor = ColVal(i, j)
    End With
    
End Sub

Private Sub BorderInitialize()
    Dim i As Long
    Dim j As Long
    Dim istr As String
    Dim jstr As String
    Dim tmp As Long
    
    For i = 0 To 11
        istr = Format(i, "00")
        For j = 0 To 11
            jstr = Format(j, "00")
            With ColPltUF.Controls("Image" & istr & jstr)
'                If .BorderColor <> 0 Then
                If .BorderStyle <> 0 Then
                    .BorderColor = 0
                    .BorderStyle = 0
                End If
            End With
        Next
    Next
End Sub

Private Sub myImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub myImage_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim ijstr As String
    Dim i As Long
    Dim j As Long
    
    BorderInitialize
    
'    If myImage0 <> "" Then
'       ColPltUF.Controls(myImage0).BorderColor = vbBlack
'    End If

    With myImage
        ijstr = Replace(.Name, "Image", "")
        .BorderStyle = 1
        .BorderColor = rgbBlack
    End With

    i = Val(Left(ijstr, 2))
    j = Val(Right(ijstr, 2))

    With ColPltUF.Controls("LabalDetermine")
        .Caption = ColName(i, j)
        .Width = 300
        .AutoSize = True
    End With

    With ColPltUF.Controls("ImageDetermine")
        .BackColor = ColVal(i, j)
    End With
    
    myImage0 = myImage.Name

End Sub
