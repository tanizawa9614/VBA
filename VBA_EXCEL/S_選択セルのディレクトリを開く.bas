Attribute VB_Name = "S_�I���Z���̃f�B���N�g�����J��"
Option Explicit

Sub �I���Z���̃f�B���N�g�����J��()
    
    Dim selectedPath As String
    
    ' �I�����ꂽ�Z������łȂ����Ƃ��m�F
    If Not IsEmpty(Selection) Then
        On Error GoTo L1
        selectedPath = Trim(Selection.Value)
        On Error GoTo 0
        
        ' �t�H���_�܂��̓t�@�C�������݂��邩�`�F�b�N
        If Dir(selectedPath, vbDirectory) <> "" Or Dir(selectedPath) <> "" Then
            ' �G�N�X�v���[���[�Ńt�H���_�܂��̓t�@�C�����J��
            Shell "explorer.exe """ & selectedPath & """", vbNormalFocus
        Else
            MsgBox "�w�肳�ꂽ�p�X��������܂���"
        End If
    End If
L1:
End Sub

