VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Test1 
   Caption         =   "UserForm1"
   ClientHeight    =   3396
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3948
   OleObjectBlob   =   "Test1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Test1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents nTextBox As MSForms.TextBox
Attribute nTextBox.VB_VarHelpID = -1
Dim n

Private Sub UserForm_Initialize()
    Me.nTextBox.Value = 1
'    Me.TextBox1.Value = 5
'    Me.TextBox2.Value = 255
'    Me.TextBox3.Value = 0
'    Me.TextBox4.Value = 255
End Sub

Private Sub CommandButton1_Click()
    Call 縁文字の実行
End Sub

Private Sub CommandButton2_Click()
    Call 処理の中断
End Sub

'Private Sub TextBox1_Change()
'    Dim n0 As Long
'    Dim i As Long
'    Dim cnt As Long
'
'    n = Me.TextBox1.Value
'    If n = "" Then Exit Sub
'    If n <= 0 Or n > 5 Then Exit Sub
'
'    With Me.MultiPage1.Pages
'        n0 = .Count
'        cnt = n0
'        If n - n0 > 0 Then
'            For i = 1 To n - n0
'                .Add
'                cnt = cnt + 1
'                Call PageAdd(cnt)
'            Next
'        ElseIf n - n0 < 0 Then
'            For i = 1 To n0 - n
'                .Clear
'            Next
'        End If
'    End With
'End Sub

'Private Sub TextBox3_Change()
'    Call Image1_Color
'
'End Sub
'
'Private Sub TextBox4_Change()
'    Call Image1_Color
'
'End Sub
'
'Private Sub TextBox5_Change()
'    Call Image1_Color
'
'End Sub

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

'Sub PageAdd(cnt As Long)
'    Dim i As Long
'    Dim idx As Long
'    Dim NewObj As Object
'    Dim OrigObj As Object
'
'    With Me.MultiPage1.Pages(cnt - 1)
'        For i = 1 To 4
'            idx = 4 * cnt - 2 + (i - 1)
'            Set NewObj = .Controls.Add("Forms.TextBox.1", "TextBox" & idx, True)
'            Set OrigObj = Me.Controls("TextBox" & i + 1)
'            Call CopyObj(OrigObj, NewObj)
'            Set NewObj = .Controls.Add("Forms.Label.1", "Label" & idx, True)
'            Set OrigObj = Me.Controls("Label" & i + 1)
'            Call CopyObj(OrigObj, NewObj)
'        Next
'        Set NewObj = .Controls.Add("Forms.Label.1", "Label" & idx + 1, True)
'        Set OrigObj = Me.Controls("Label" & 6)
'        Call CopyObj(OrigObj, NewObj)
'        Set NewObj = .Controls.Add("Forms.Image.1", "Image" & cnt, True)
'        Set OrigObj = Me.Controls("Image" & 1)
'        Call CopyObj(OrigObj, NewObj)
'    End With
'End Sub

Private Sub CopyObj(Obj1 As Object, Obj2 As Object)
    On Error Resume Next
    Obj2.Caption = Obj1.Caption
    Obj2.Text = Obj1.Text
    Obj2.Left = Obj1.Left
    Obj2.Top = Obj1.Top
    Obj2.Height = Obj1.Height
    Obj2.Width = Obj1.Width
    On Error GoTo 0
End Sub


'*****************************

Private Sub nTextBox_Change()
    ' nの値が変更された場合、ページ数を変更する
    Dim n As Integer
    n = Val(nTextBox.Text)
    
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
            NewTextBox.Top = 10 + (j - 1) * 30
            NewTextBox.Left = 10
            NewTextBox.Width = 50
            
            ' テキストボックスのChangeイベントを処理するためのクラスを作成し、イベントを処理する
            Dim eventHandler As New TextBoxEventHandler
            Set eventHandler.TextBox = NewTextBox
            
            Dim NewLabel As MSForms.Label
            Set NewLabel = newPage.Controls.Add("Forms.Label.1", "Label" & j)
            NewTextBox.Top = 10 + (j - 1) * 30
            NewTextBox.Left = 10
            NewTextBox.Width = 50
        Next j
    Next i
End Sub


