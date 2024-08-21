Attribute VB_Name = "S_Auto_Open"
Option Explicit

Option Explicit
Const ControlCaption As String = "Addin"

Public Sub Auto_Open()
'アドインとして読み込んだときに実行
  With Application.CommandBars("Standard")
    On Error Resume Next
    .Controls(ControlCaption).Delete
    On Error GoTo 0
    With .Controls.Add(Type:=msoControlButton) '[アドイン]タブにボタン追加
      .Caption = ControlCaption
      .Style = msoButtonIconAndCaption
      .FaceId = 65 'アイコン画像の設定
      .OnAction = "Sample"
    End With
  End With
End Sub

Public Sub Auto_Close()
'アドインの読み込み解除したときに実行
  On Error Resume Next
  Application.CommandBars("Standard").Controls(ControlCaption).Delete
  On Error GoTo 0
End Sub

Public Sub Sample()
'ボタンをクリックしたときに呼ばれるプロシージャー
  MsgBox "OK", vbSystemModal
End Sub
