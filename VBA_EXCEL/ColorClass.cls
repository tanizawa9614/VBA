VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColorClass"
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
    
'    BorderInitialize
    
    With myImage
'        .BorderColor = RGB(255, 255, 255)
        ijstr = Replace(.Name, "Image", "")
    End With
    
    i = Val(Left(ijstr, 2))
    j = Val(Right(ijstr, 2))
    
    With ColUserForm.Controls("LabalCurrent")
        .Caption = ColName(i, j)
        .Width = 300
        .AutoSize = True
    End With
    
    With ColUserForm.Controls("ImageCurrent")
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
            With ColUserForm.Controls("Image" & istr & jstr)
                tmp = .BorderStyle
                If .BorderStyle = 0 Then
                    .BorderStyle = 1
                End If
            End With
        Next
    Next
End Sub

Private Sub myImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim ijstr As String
    Dim i As Long
    Dim j As Long

    BorderInitialize

    With myImage
        ijstr = Replace(.Name, "Image", "")
        .BorderStyle = 0
    End With

    i = Val(Left(ijstr, 2))
    j = Val(Right(ijstr, 2))

    With ColUserForm.Controls("LabalDetermine")
        .Caption = ColName(i, j)
        .Width = 300
        .AutoSize = True
    End With

    With ColUserForm.Controls("ImageDetermine")
        .BackColor = ColVal(i, j)
    End With
End Sub
