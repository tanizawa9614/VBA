Attribute VB_Name = "S_Auto_Open"
Option Explicit

Option Explicit
Const ControlCaption As String = "Addin"

Public Sub Auto_Open()
'�A�h�C���Ƃ��ēǂݍ��񂾂Ƃ��Ɏ��s
  With Application.CommandBars("Standard")
    On Error Resume Next
    .Controls(ControlCaption).Delete
    On Error GoTo 0
    With .Controls.Add(Type:=msoControlButton) '[�A�h�C��]�^�u�Ƀ{�^���ǉ�
      .Caption = ControlCaption
      .Style = msoButtonIconAndCaption
      .FaceId = 65 '�A�C�R���摜�̐ݒ�
      .OnAction = "Sample"
    End With
  End With
End Sub

Public Sub Auto_Close()
'�A�h�C���̓ǂݍ��݉��������Ƃ��Ɏ��s
  On Error Resume Next
  Application.CommandBars("Standard").Controls(ControlCaption).Delete
  On Error GoTo 0
End Sub

Public Sub Sample()
'�{�^�����N���b�N�����Ƃ��ɌĂ΂��v���V�[�W���[
  MsgBox "OK", vbSystemModal
End Sub
