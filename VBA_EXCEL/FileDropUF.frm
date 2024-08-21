VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileDropUF 
   Caption         =   "ファイルをここにドロップしてください"
   ClientHeight    =   700
   ClientLeft      =   -140
   ClientTop       =   -600
   ClientWidth     =   480
   OleObjectBlob   =   "FileDropUF.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FileDropUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    With ListView1
    
        ''''プロパティの設定
        .FullRowSelect = True           '行全体の選択
'        .Gridlines = True               '行列グリッド線の表示
'        .View = lvwReport               '表示形式
        .OLEDropMode = ccOLEDropManual  'ファイルドロップ処理

        ''''列見出しの名前・列幅の設定
'        .ColumnHeaders.Add , "key1", "ここにファイルをドロップしてください", 450, lvwColumnLeft
        .Width = 1000#
        .Height = 1000
        .Left = -10
        .BackColor = RGB(150, 150, 150)
    End With
    Me.Caption = "ファイルをドロップすると参考文献用のテキストが出力されます"
    Me.Width = 400
    Me.Height = 300
End Sub


Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim i As Long
    Dim FileCount As Long
    Dim Arr() As String
    
    
    With ListView1
        
        'ファイル数の取得（複数ファイルを同時にドラッグ＆ドロップした時用）
        FileCount = Data.Files.Count
        ReDim Arr(1 To FileCount, 1 To 1)
                
'        ドラッグ& ドロップしたファイルパスを順にリスト化
        For i = 1 To FileCount
            Arr(i, 1) = Data.Files(i)
'            .ListItems.Add = Data.Files(i)
        Next i
    End With
    Unload Me
    テキストファイル読み込み_Uf (Arr)
End Sub
