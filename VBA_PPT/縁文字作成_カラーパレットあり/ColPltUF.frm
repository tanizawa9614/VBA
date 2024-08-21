VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColPltUF 
   Caption         =   "UserForm1"
   ClientHeight    =   4236
   ClientLeft      =   72
   ClientTop       =   288
   ClientWidth     =   4584
   OleObjectBlob   =   "ColPltUF.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ColPltUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm
Option Explicit
Private ImArr(11, 11) As New ImageEvents
Private WithEvents OKButton As MSForms.CommandButton
Attribute OKButton.VB_VarHelpID = -1
Private WithEvents CancelButton As MSForms.CommandButton
Attribute CancelButton.VB_VarHelpID = -1
Dim LabalCurrent As MSForms.Label
Dim LabalDetermine As MSForms.Label
Dim ImageCurrent As MSForms.Image
Dim ImageDetermine As MSForms.Image
Dim TextCurrent As MSForms.Label
Dim TextDetermine As MSForms.Label

Private Sub UserForm_Initialize()
    Dim i As Long
    Dim j As Long
    Dim NewImage As MSForms.Image
    Dim x0 As Double
    Dim y0 As Double
    Dim x1 As Double
    Dim y1 As Double
    
    SizeMe = Me.Height / 18
    Me.StartUpPosition = 0
    Me.Left = EdgeCharUf.Left
    Me.Top = EdgeCharUf.Top
    Me.Width = Me.Height + 100
    Me.Caption = "色の設定"
    Me.BackColor = RGB(234, 234, 234)
    Me.BorderStyle = 0
    
    Set OKButton = Me.Controls.Add("Forms.CommandButton.1", "OK", True)
    With OKButton
        .Width = Me.Width / 4
        .Height = Me.Height / 8
        .Top = 10
        .Left = Me.Width * 2 / 3
        .Caption = "OK"
        .BackColor = rgbWhiteSmoke
    End With
    
    Set CancelButton = Me.Controls.Add("Forms.CommandButton.1", "Cancel", True)
    With CancelButton
        .Width = Me.Width / 4
        .Height = Me.Height / 8
        .Top = OKButton.Top + OKButton.Height + 10
        .Left = Me.Width * 2 / 3
        .Caption = "キャンセル [C]"
        .BackColor = rgbWhiteSmoke
    End With
    
    MakeColorArray
    Dim istr As String
    Dim jstr As String
    For i = 0 To 11
        istr = Format(i, "00")
        For j = 0 To 11
            jstr = Format(j, "00")
            Set NewImage = Me.Controls.Add("Forms.Image.1", "Image" & istr & jstr, True)
            With NewImage
                .BackColor = ColVal(i, j)
                .Top = SizeMe * i + 10 + i * 1.5
                .Left = SizeMe * j + 20 + j * 1.5
                .Width = SizeMe
                .Height = SizeMe
                .BorderColor = 0
                .BorderStyle = 0
                If i = 0 And j = 0 Then
                    x0 = .Left
                    y0 = .Top
                ElseIf i = 11 And j = 11 Then
                    x1 = .Left + .Width
                    y1 = .Top + .Height
                End If
            End With
            ImArr(i, j).SetImage NewImage
        Next
    Next
    
    Set ImageCurrent = Me.Controls.Add("Forms.Image.1", "ImageCurrent")
    With ImageCurrent
        .Top = Me.Height - 70
        .Left = Me.Width - 120
        .Height = SizeMe * 2
        .Width = .Height * 2
        .BackColor = ColVal(0, 0)
    End With
    
    Set LabalCurrent = Me.Controls.Add("Forms.Label.1", "LabalCurrent")
    With LabalCurrent
        .Top = ImageCurrent.Top - 12
        .Left = ImageCurrent.Left
        .Width = 300
        .Font.size = 11
        .Caption = ColName(0, 0)
        .AutoSize = True
        .BackStyle = 0
    End With
    
    Set TextCurrent = Me.Controls.Add("Forms.Label.1", "LabalCurrent")
    With TextCurrent
        .Top = LabalCurrent.Top - 12
        .Left = LabalCurrent.Left
        .Width = 300
        .Font.size = 11
        .Caption = "マウス上の色 : "
        .AutoSize = True
        .BackStyle = 0
    End With
    
    Set ImageDetermine = Me.Controls.Add("Forms.Image.1", "ImageDetermine")
    With ImageDetermine
        .Top = TextCurrent.Top - 35
        .Left = TextCurrent.Left
        .Width = ImageCurrent.Width
        .Height = ImageCurrent.Height
        .BackColor = ColVal(0, 0)
    End With
    
    Set LabalDetermine = Me.Controls.Add("Forms.Label.1", "LabalDetermine")
    With LabalDetermine
        .Top = ImageDetermine.Top - 12
        .Left = ImageDetermine.Left
        .Width = LabalCurrent.Width
        .Font.size = LabalCurrent.Font.size
        .Caption = ColName(0, 0)
        .AutoSize = True
        .BackStyle = 0
    End With
    
    Set TextDetermine = Me.Controls.Add("Forms.Label.1", "LabalCurrent")
    With TextDetermine
        .Top = LabalDetermine.Top - 12
        .Left = LabalDetermine.Left
        .Width = 300
        .Font.size = 11
        .Caption = "選択中の色  : "
        .AutoSize = True
        .BackStyle = 0
    End With
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If IsInObj(OKButton, X, Y) = False Then Call ColorGray(OKButton)
    If IsInObj(CancelButton, X, Y) = False Then Call ColorGray(CancelButton)
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub CancelButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ColorBlue(CancelButton)
End Sub

Private Sub OKButton_Click()
    AnsCol = ImageDetermine.BackColor
    Me.Hide
    
    ' UserForm1のコマンドボタンに操作を返す前にrgb値を書き換え
    Dim R As Long
    Dim G As Long
    Dim B As Long
    Dim arr
    B = AnsCol \ (256 ^ 2)
    G = (AnsCol - 256 ^ 2 * B) \ 256
    R = AnsCol - 256 ^ 2 * B - 256 * G
    arr = Array(R, G, B)
    
    Dim cPage As Long, i As Long
    cPage = EdgeCharUf.MultiPage1.Value
    With EdgeCharUf.MultiPage1.Pages(cPage)
        For i = 2 To 4
            .Controls("TextBox" & i).Value = arr(i - 2)
        Next
    End With
End Sub

Private Sub OKButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ColorBlue(OKButton)
End Sub

