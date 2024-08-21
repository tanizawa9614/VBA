VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Uf_文字の大きさ一括変更 
   Caption         =   "文字のサイズを入力してください"
   ClientHeight    =   1668
   ClientLeft      =   84
   ClientTop       =   372
   ClientWidth     =   3012
   OleObjectBlob   =   "Uf_文字の大きさ一括変更.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Uf_文字の大きさ一括変更"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim tmpsize As Double
'    On Error Resume Next
    tmpsize = Val(TextBox1.Value)
    If Err.Number > 0 Or tmpsize <= 0 Then
        Unload Me
        Exit Sub
    End If
    On Error GoTo 0
    ChangeStringSize (tmpsize)
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    UFPositionCenter
End Sub

Private Sub UFPositionCenter()
    '**ユーザーフォームを親ウィンドウの中央に表示する
    
    '**変数(T=Top,L=Left,W=Width,H=Height,AW=ActiveWindow,UF=UserForm)
    Dim T_AW As Long, L_AW As Long, W_AW As Long, H_AW As Long
    Dim T_UF As Long, L_UF As Long, W_UF As Long, H_UF As Long
    
    '**親ウィンドウの位置とサイズを取得
    With ActiveWindow
    T_AW = .Top
    L_AW = .Left
    W_AW = .Width
    H_AW = .Height
    End With
    
    '**UFのサイズを取得
    W_UF = Me.Width
    H_UF = Me.Height
    
    '**UFの表示位置を計算
    T_UF = T_AW + ((H_AW - H_UF) / 2)
    L_UF = L_AW + ((W_AW - W_UF) / 2)
    
    '**UFの表示位置を設定
    Me.StartUpPosition = 0
    '**Top,Left指定時に必須(ないとLeftがずれる)
    Me.Top = T_UF
    Me.Left = L_UF
End Sub
