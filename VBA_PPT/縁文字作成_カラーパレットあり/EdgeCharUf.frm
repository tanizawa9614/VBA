VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EdgeCharUf 
   Caption         =   "UserForm3"
   ClientHeight    =   4284
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4296
   OleObjectBlob   =   "EdgeCharUf.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "EdgeCharUf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'UserForm
Option Explicit
Private WithEvents nTextBox As MSForms.TextBox
Attribute nTextBox.VB_VarHelpID = -1
Private TbArr() As New TextBoxEvents
Private CbArr() As New TextBoxEvents
Private LWArr() As New TextBoxEvents
Private CBColPalt As MSForms.CommandButton
Dim h As Double
Dim firstflg As Boolean
Dim nmax As Long

Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ColorBlue(CommandButton1)
End Sub

Private Sub CommandButton2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ColorBlue(CommandButton2)
End Sub

Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim CBtmp As MSForms.CommandButton
    Dim cPage As Long
    With Me.MultiPage1
        cPage = .Value
        Set CBtmp = .Pages(cPage).Controls("CBColPalt")
        If IsInObj(CBtmp, X, Y) = False Then Call ColorGray(CBtmp)
    End With
End Sub

Private Sub nTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim ActivePage As Long
    ActivePage = Me.MultiPage1.Value
    
    With nTextBox
        Select Case KeyCode
        Case vbKeyUp
            KeyCode = 0
            If .Value >= nmax Then Exit Sub
            .Value = Val(.Value) + 1
        Case vbKeyDown
            KeyCode = 0
            If .Value <= 1 Then Exit Sub
            .Value = Val(.Value) - 1
        Case vbKeySpace, vbKeyTab
            KeyCode = 0
            With Me.MultiPage1
                .Value = ActivePage
                .Pages(ActivePage).Controls("TextBox1").SetFocus
            End With
        End Select
    End With
End Sub

Private Sub UserForm_Initialize()
    Dim n As Long
    n = 2 ' ページ数の初期値
    nmax = 15
    
    Set nTextBox = Me.Controls.Add("Forms.TextBox.1", "nTextBox")
    With nTextBox
        .Top = 20
        .Left = 50
        .Width = 50
        .Height = Const_h
        .TabIndex = 3
        firstflg = 1
        .Value = n
        firstflg = 0
'        .EnterFieldBehavior = fmEnterFieldBehaviorSelectAll
        .SetFocus
        .TextAlign = fmTextAlignRight
        .IMEMode = fmIMEModeDisable
        .TabKeyBehavior = True
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    With Me.LabelN
        .Accelerator = "N"
        .TabIndex = 2
        .Top = nTextBox.Top + 2
    End With
    
    Me.Caption = "縁文字の作成"
    
    With Me.CommandButton1
        .Default = True
        .TabIndex = 0
    End With
    With Me.CommandButton2
        .Default = False
        .Cancel = True
        .TabIndex = 1
        .Accelerator = "C"
    End With
End Sub

Private Sub nTextBox_Change()
    Dim n As Long
    Dim i As Long
    Dim n0 As Long
    
    n = Val(nTextBox.Text)
    n0 = MultiPage1.Pages.Count
    
    If n <= 0 Or n > nmax Then
        n = nmax
        If n <= 0 Then n = 1
    End If
        
    If n > n0 Then ' 初期よりページ数が多い
        If firstflg Then Call UpdatePages(1, n0)
        For i = n0 + 1 To n
            MultiPage1.Pages.Add
        Next i
        Call UpdatePages(n0 + 1, n)
    ElseIf n < n0 Then  ' 初期よりページ数が少ない
        If firstflg Then Call UpdatePages(1, n)
        For i = n0 - 1 To n Step -1
            MultiPage1.Pages.Remove i
        Next i
        ReDim Preserve TbArr(2, n - 1)
        ReDim Preserve LWArr(n - 1)
        With MultiPage1.Pages(0).Controls("TextBox1")
            .Value = .Value + 0.1
            .Value = .Value - 0.1
        End With
    ElseIf n = n0 And firstflg Then  ' 初期と同じ
        Call UpdatePages(1, n)
    End If
    
    For i = 1 To MultiPage1.Pages.Count
        MultiPage1.Pages(i - 1).Caption = "縁 " & i
    Next i
    Controls("nTextBox").SetFocus
End Sub

Private Sub UpdatePages(n0 As Long, n As Long)
    Randomize
    Dim i As Long
    Dim cnt As Long
    Dim pPage As MSForms.Page
    
    ReDim Preserve TbArr(2, n - 1)
    ReDim Preserve CbArr(n - 1)
    ReDim Preserve LWArr(n - 1)
    
    For i = n0 - 1 To n - 1
        
        Set pPage = MultiPage1.Pages(i)
        MultiPage1.Value = i
        
        Dim PreView As MSForms.Label
        Set PreView = pPage.Controls.Add("Forms.Label.1", "PreView")
        PreView.Top = 25
        PreView.Left = 110
        PreView.Width = 100
        PreView.Caption = "カラープレビュー :"
        PreView.AutoSize = True
            
        Dim Image As MSForms.Image
        Set Image = pPage.Controls.Add("Forms.Image.1", "Image")
        Image.Top = PreView.Top + 15
        Image.Left = PreView.Left + 10
        Image.Width = 60
        Image.Height = Image.Width
        
        Dim j As Long
        For j = 0 To 3
            
            Call Common_Label(pPage, j)
            
            Dim tb As MSForms.TextBox
            Set tb = pPage.Controls.Add("Forms.TextBox.1", "TextBox" & j + 1, True)
            Call Common_TextBox_Font(tb, j)
            
            If j = 0 Then
                LWArr(i).SetLWTextBox tb
            Else
                TbArr(j - 1, i).SetTextBox tb
            End If
            Dim Lwtmp As Double, fPageLW As Double
            If j = 0 Then ' 線の太さの指定
                Lwtmp = 3 * (i + 1)
                If i <> 0 Then '一ページ目以外は前のページより大きい値にする
                    fPageLW = Val(MultiPage1.Pages(i - 1).Controls("TextBox1")) + 3
                    If fPageLW > Lwtmp Then Lwtmp = fPageLW
                End If
                tb.Text = Lwtmp
            Else
                tb.Text = Int(Rnd * (255 - 0 + 1) + 0) 'RGBの初期値
            End If
            cnt = cnt + 1
        Next j
        
        ' コマンドボタン（カラーパレット表示用コマンドボタン）の追加
        Set CBColPalt = pPage.Controls.Add("Forms.CommandButton.1", "CBColPalt")
        With CBColPalt
            .Top = 5
            .Left = 3
            .Height = 20
            .Width = 97
            .Caption = "カラーパレット"
        End With
        CbArr(i).SetCommandButton CBColPalt
    Next i
End Sub

Private Sub Common_TextBox_Font(NewTextBox As MSForms.TextBox, ByVal j As Long)
    h = Const_h
    With NewTextBox
        .Top = 55 + (j - 1) * h + 8 * (j - 1)
        .Left = 65
        .Width = 35
        .Height = h
        .Value = 25
        .TextAlign = fmTextAlignRight
        .IMEMode = fmIMEModeDisable
        .TabKeyBehavior = False
    End With
End Sub

Private Sub Common_Label(p As MSForms.Page, j As Long)
    Dim NewLabel As MSForms.Label
    Set NewLabel = p.Controls.Add("Forms.Label.1", "Label" & j)
    h = Const_h
    
    Select Case j
        Case 0
            NewLabel.Caption = "線の太さ [W] :"
            NewLabel.Accelerator = "W"
        Case 1
            NewLabel.Caption = "   R      [R] : "
            NewLabel.Accelerator = "R"
        Case 2
            NewLabel.Caption = "   G      [G] : "
            NewLabel.Accelerator = "G"
        Case 3
            NewLabel.Caption = "   B      [B] : "
            NewLabel.Accelerator = "B"
    End Select
    NewLabel.Top = 55 + (j - 1) * h + 8 * (j - 1)
    NewLabel.Left = 6
    NewLabel.Width = 100
    NewLabel.Height = h
    NewLabel.AutoSize = True
    NewLabel.TextAlign = fmTextAlignRight
End Sub

Private Function Const_h() As Double
    Const_h = 15
End Function

Private Sub CommandButton1_Click()
    Make_LineWidth_rgbArr
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    CancelFlg = True
    Run縁文字の解除
    Unload Me
End Sub

Public Sub Make_LineWidth_rgbArr()
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim rgbArr(0 To 2)
    Errflg = False
    
    With Me.MultiPage1
        n = .Pages.Count
        ReDim rgbVal(n - 1)
        ReDim LineWidth(n - 1)
        For i = 0 To n - 1
            On Error GoTo ErrHdl
            For j = 0 To 2
                rgbArr(j) = Val(.Pages(i).Controls("TextBox" & j + 2))
            Next
            rgbVal(i) = RGB(rgbArr(0), rgbArr(1), rgbArr(2))
            LineWidth(i) = Val(.Pages(i).Controls("TextBox" & 1))
            On Error GoTo 0
        Next
    End With
    Exit Sub
ErrHdl:
    Errflg = True
End Sub

'コマンドボタンの塗りつぶしを元に戻す
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If IsInObj(CommandButton1, X, Y) = False Then Call ColorGray(CommandButton1)
    If IsInObj(CommandButton2, X, Y) = False Then Call ColorGray(CommandButton2)
    Dim CBtmp As MSForms.CommandButton
    Dim x2 As Single, y2 As Single, cPage As Long
    With Me.MultiPage1
        cPage = .Value
        Set CBtmp = .Pages(cPage).Controls("CBColPalt")
        x2 = X - .Left
        y2 = Y - .Top - 18
        If IsInObj(CBtmp, x2, y2) = False Then Call ColorGray(CBtmp)
    End With
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' ×ボタンが押された場合
    If CloseMode = vbFormControlMenu Then
        Run縁文字の解除
        Unload Me
        CancelFlg = True
    End If
End Sub
