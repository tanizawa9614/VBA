VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColUserForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4236
   ClientLeft      =   72
   ClientTop       =   288
   ClientWidth     =   4584
   OleObjectBlob   =   "ColUserForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ColUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm
Option Explicit
Private ImArr(11, 11) As New ColorClass
Private WithEvents OKButton As MSForms.CommandButton
Attribute OKButton.VB_VarHelpID = -1
Private WithEvents CancelButton As MSForms.CommandButton
Attribute CancelButton.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    Dim i As Long
    Dim j As Long
    Dim NewImage As MSForms.Image
    Dim LabalCurrent As MSForms.Label
    Dim LabalDetermine As MSForms.Label
    Dim ImageCurrent As MSForms.Image
    Dim ImageDetermine As MSForms.Image
    Dim TextCurrent As MSForms.Label
    Dim TextDetermine As MSForms.Label
    
    SizeMe = Me.Height / 15
    Me.Width = Me.Height + 100
    Me.Caption = "カラーパレット : 色を選択してください"
'    Me.BackColor = ColVal(0, 1)
    
    Set OKButton = Me.Controls.Add("Forms.CommandButton.1", "OK", True)
    With OKButton
        .Width = Me.Width / 4
        .Height = Me.Height / 8
        .Top = 10
        .Left = Me.Width * 2 / 3
        .Caption = "OK"
    End With
    
    Set CancelButton = Me.Controls.Add("Forms.CommandButton.1", "Cancel", True)
    With CancelButton
        .Width = Me.Width / 4
        .Height = Me.Height / 8
        .Top = OKButton.Top + OKButton.Height + 10
        .Left = Me.Width * 2 / 3
        .Caption = "キャンセル [C]"
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
                .Top = SizeMe * i + 10
                .Left = SizeMe * j + 20
                .Width = SizeMe
                .Height = SizeMe
                .BorderColor = 0
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

Private Sub OKButton_Click()
    Dim getCol As String
    Dim i As Long
    Dim j As Long
    Dim AnsColor As Long
    
    getCol = ColorLabel.Caption
    
    For i = 0 To 11
        For j = 0 To 11
            If getCol = ColName(i, j) Then GoTo L1
        Next
    Next
L1:
    AnsColor = ColVal(i, j)
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub
