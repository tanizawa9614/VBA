Attribute VB_Name = "S_選択セルのディレクトリを開く"
Option Explicit

Sub 選択セルのディレクトリを開く()
    
    Dim selectedPath As String
    
    ' 選択されたセルが空でないことを確認
    If Not IsEmpty(Selection) Then
        On Error GoTo L1
        selectedPath = Trim(Selection.Value)
        On Error GoTo 0
        
        ' フォルダまたはファイルが存在するかチェック
        If Dir(selectedPath, vbDirectory) <> "" Or Dir(selectedPath) <> "" Then
            ' エクスプローラーでフォルダまたはファイルを開く
            Shell "explorer.exe """ & selectedPath & """", vbNormalFocus
        Else
            MsgBox "指定されたパスが見つかりません"
        End If
    End If
L1:
End Sub

