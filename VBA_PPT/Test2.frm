VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Test2 
   Caption         =   "UserForm2"
   ClientHeight    =   3816
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3504
   OleObjectBlob   =   "Test2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Test2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents nTextBox As MSForms.TextBox
Attribute nTextBox.VB_VarHelpID = -1
Dim n

Private Sub UserForm_Initialize()
    Set nTextBox = Me.Controls.Add("Forms.TextBox.1", "nTextBox")
    nTextBox.Top = 20
    nTextBox.Left = 50
    nTextBox.Width = 30
    nTextBox.Value = 1
    nTextBox.SetFocus
End Sub

Private Sub CommandButton1_Click()
    Call 縁文字の実行
End Sub

Private Sub CommandButton2_Click()
    Call 処理の中断
End Sub

Sub Image1_Color(arr)
    Dim R, G, B
    R = arr(1)
    G = arr(2)
    B = arr(3)
    If R = "" Then Exit Sub
    If G = "" Then Exit Sub
    If B = "" Then Exit Sub
    Me.Image1.BackColor = RGB(Val(R), Val(G), Val(B))

End Sub

Private Sub nTextBox_Change()
    ' nの値が変更された場合、ページ数を変更する
    Dim n As Integer
    n = Val(nTextBox.Text)
    
    If n <= 0 Then
        Exit Sub
    ElseIf n >= 6 Then
        n = 5
        On Error Resume Next
        If Me.Controls("ErrMsg") Is Nothing Then
            Dim ErrMsg As MSForms.Label
            Set ErrMsg = Me.Controls.Add("Forms.Label.1", "ErrMsg")
            ErrMsg.Top = 35
            ErrMsg.Left = 30
            ErrMsg.Height = 18
            ErrMsg.Width = 150
            ErrMsg.ForeColor = RGB(255, 0, 0)
            ErrMsg.Caption = "5以下の数値を" & vbCrLf & "代入してください"
            ErrMsg.AutoSize = True
        Else
            Me.Controls("ErrMsg").Visible = True
        End If
    Else
        On Error Resume Next
        Me.Controls("ErrMsg").Visible = False
        On Error GoTo 0
    End If
    
    ' ページ数を変更する前に、最初のページ以外の既存のページを削除する
    Do While Me.MultiPage1.Pages.Count > 0
        Me.MultiPage1.Pages.Remove (0)
    Loop
    
    ' 新しいページを追加する
    Dim i As Integer
    For i = 1 To n
        Dim newPage As MSForms.Page
        Set newPage = Me.MultiPage1.Pages.Add("Page" & i, "Page" & i)
        
        ' 各ページに4つのテキストボックスを追加する
        Dim j As Integer
        For j = 1 To 4
            Dim NewTextBox As MSForms.TextBox
            Set NewTextBox = newPage.Controls.Add("Forms.TextBox.1", "TextBox" & j)
            NewTextBox.Top = 6 + (j - 1) * 20
            NewTextBox.Left = 45
            NewTextBox.Width = 40
            NewTextBox.Height = 18
            NewTextBox.Value = 25
                        
            ' テキストボックスのChangeイベントを処理するためのクラスを作成し、イベントを処理する
            Dim eventHandler As New TextBoxEventHandler
            Set eventHandler.TextBox = NewTextBox
            
            Dim NewLabel As MSForms.Label
            Set NewLabel = newPage.Controls.Add("Forms.Label.1", "Label" & j)
            NewLabel.Top = 8 + (j - 1) * 20
            NewLabel.Left = 6
            NewLabel.Width = 40
            NewLabel.Height = 18
            Select Case j
                Case 1
                    NewLabel.Caption = "線の太さ :"
                Case 2
                    NewLabel.Caption = "   R       : "
                Case 3
                    NewLabel.Caption = "   G       : "
                Case 4
                    NewLabel.Caption = "   B       : "
            End Select
            
        Next j
        
        Dim PreView As MSForms.Label
        Set PreView = newPage.Controls.Add("Forms.Label.1", "PreView")
        PreView.Top = 20
        PreView.Left = 100
        PreView.Width = 100
        PreView.Caption = "プレビュー :"
        PreView.AutoSize = True
        
        Dim Image As MSForms.Image
        Set Image = newPage.Controls.Add("Forms.Image.1", "Image")
        Image.Top = 30
        Image.Left = 100
        Image.Width = 50
        Image.Height = 50
    Next i
    
    
End Sub



