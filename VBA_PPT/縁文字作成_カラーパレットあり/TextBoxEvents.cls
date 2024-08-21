VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' TextBoxEvents クラスのコード
Dim WithEvents myTextBox As MSForms.TextBox
Attribute myTextBox.VB_VarHelpID = -1
Dim WithEvents myLWTextBox As MSForms.TextBox
Attribute myLWTextBox.VB_VarHelpID = -1
Dim WithEvents myCommandButton As MSForms.CommandButton
Attribute myCommandButton.VB_VarHelpID = -1
Dim LWCancelflg As Boolean

Public Sub SetTextBox(NewTextBox As MSForms.TextBox)
    Set myTextBox = NewTextBox
End Sub
Public Sub SetLWTextBox(NewTextBox As MSForms.TextBox)
    Set myLWTextBox = NewTextBox
'    LWCancelflg = False
End Sub
Public Sub SetCommandButton(NewCommandButton As MSForms.CommandButton)
    Set myCommandButton = NewCommandButton
End Sub

Private Sub myCommandButton_Click()
    EdgeCharUf.Hide
    ' UserForm2: カラーパレットの表示
    CallColPalette 'ColorPalette用の標準モジュール
    
    With EdgeCharUf
        .StartUpPosition = 0
        .Left = ColPltUF.Left
        .Top = ColPltUF.Top
        Unload ColPltUF
        .Show
    End With
End Sub

Private Sub myCommandButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ColorBlue(myCommandButton)
End Sub

Private Sub myLWTextBox_Change()
    Dim n As Long, ipage As Long
    Dim beforelw As Double, currentlw As Double, afterlw As Double
    
    If LWCancelflg = True Then Exit Sub 'イベントの連鎖を防ぐ
    LWCancelflg = True
    
    With myLWTextBox
        If Val(.Value) < 0 Then .Value = 1
        If IsNumeric(.Value) = False Then .Value = 1
        With EdgeCharUf.MultiPage1 ' LWの調整（前の縁よりは太く）
            n = .Pages.Count
            ipage = .Value
            currentlw = Val(myLWTextBox.Value)
            If ipage <> 0 Then
                beforelw = Val(.Pages(ipage - 1).Controls("TextBox1").Value)
                If currentlw <= beforelw Then myLWTextBox.Value = beforelw + 1
            End If
            If ipage + 1 <= n - 1 Then
                On Error Resume Next
                afterlw = Val(.Pages(ipage + 1).Controls("TextBox1").Value)
                If afterlw <> 0 And currentlw >= afterlw Then
                    .Value = ipage + 1
                    .Pages(ipage + 1).Controls("TextBox1").Value = currentlw + 1
                    .Value = ipage
                End If
                On Error GoTo 0
            End If
        End With
        Run縁文字の解除
        EdgeCharUf.Make_LineWidth_rgbArr
        If Errflg = False Then Run縁文字の作成
    End With
    LWCancelflg = False 'Falseに戻してイベントが発生するように
myHdl:
End Sub

Private Sub myLWTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call Common_TextBox_KeyDown(myLWTextBox, KeyCode)
End Sub

Private Sub myTextBox_Change()
    Dim i As Integer
    Dim rgbArr(1 To 3)
    Dim tb As MSForms.TextBox
    Dim tbName As String
    
    With myTextBox
        If Val(.Value) > 255 Then .Value = 255
        If Val(.Value) < 0 Then .Value = 0
        If IsNumeric(.Value) = False Then .Value = 0
        
        Run縁文字の解除
        For i = 2 To 4
            tbName = "TextBox" & i
            On Error GoTo myHdl
            Set tb = .Parent.Controls(tbName)
            rgbArr(i - 1) = Val(tb.Text)
            On Error GoTo 0
        Next i
        .Parent.Controls("Image").BackColor = RGB(rgbArr(1), rgbArr(2), rgbArr(3))
        EdgeCharUf.Make_LineWidth_rgbArr
        If Errflg = False Then Run縁文字の作成
        Errflg = True
    End With
myHdl:
End Sub

Private Sub myTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call Common_TextBox_KeyDown(myTextBox, KeyCode)
End Sub

Private Sub Common_TextBox_KeyDown(Obj As MSForms.TextBox, ByVal KeyCode As MSForms.ReturnInteger)
'LW，RGBテキストボックスに共通のイベント
    With Obj
        Select Case KeyCode
        Case vbKeyUp
            KeyCode = 0
            If .Value > 255 Or .Value < 0 Then Exit Sub
            .Value = Val(.Value) + 1
        Case vbKeyDown
            KeyCode = 0
            If .Value > 255 Or .Value < 0 Then Exit Sub
            .Value = Val(.Value) - 1
        Case vbKeySpace, vbKeyTab
            .SelStart = 0
            .SelLength = Len(.Text)
'            KeyCode = 0
        End Select
    End With
End Sub
