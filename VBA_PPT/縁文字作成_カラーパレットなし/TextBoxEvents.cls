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

Sub SetTextBox(NewTextBox As MSForms.TextBox)
    Set myTextBox = NewTextBox
End Sub

Private Sub myTextBox_Change()
    Dim i As Integer
    Dim rgbArr(1 To 3)
    Dim tb As MSForms.TextBox
    Dim tbName As String
    
    With myTextBox
        If Val(.Value) > 255 Then .Value = 255
        If Val(.Value) < 0 Then .Value = 0
        
        Run縁文字の解除
        For i = 2 To 4
            tbName = "TextBox" & i
            On Error GoTo myHdl
            Set tb = .Parent.Controls(tbName)
            rgbArr(i - 1) = Val(tb.Text)
            If IsNumeric(tb.Text) = False Then Exit Sub
            On Error GoTo 0
        Next i
        .Parent.Controls("Image").BackColor = RGB(rgbArr(1), rgbArr(2), rgbArr(3))
        BuildUf.Make_LineWidth_rgbArr
        If Errflg = False Then Run縁文字の作成
    End With
myHdl:
End Sub

Private Sub myTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With myTextBox
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
