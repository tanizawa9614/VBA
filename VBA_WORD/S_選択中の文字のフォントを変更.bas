Attribute VB_Name = "S_�I�𒆂̕����̃t�H���g��ύX"
Option Explicit

Sub �I�𒆂̕����̃t�H���g��ύX()
    Const font1 As String = "�l�r �S�V�b�N"
    Const font2 As String = "Times New Roman"
    
     
    ' �e�L�X�g���I������Ă��邱�Ƃ��m�F
    If Selection.Type = wdSelectionNormal Then
        ' �I�������e�L�X�g�̃t�H���g��ύX
        Selection.Font.NameFarEast = font1
        Selection.Font.Name = font2
    End If
End Sub

