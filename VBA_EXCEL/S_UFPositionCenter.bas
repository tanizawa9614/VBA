Attribute VB_Name = "S_UFPositionCenter"
Option Explicit

Private Sub UFPositionCenter()
    '**���[�U�[�t�H�[����e�E�B���h�E�̒����ɕ\������
    
    '**�ϐ�(T=Top,L=Left,W=Width,H=Height,AW=ActiveWindow,UF=UserForm)
    Dim T_AW As Long, L_AW As Long, W_AW As Long, H_AW As Long
    Dim T_UF As Long, L_UF As Long, W_UF As Long, H_UF As Long
    
    '**�e�E�B���h�E�̈ʒu�ƃT�C�Y���擾
    With ActiveWindow
        T_AW = .Top
        L_AW = .Left
        W_AW = .Width
        H_AW = .Height
    End With
    
    '**UF�̃T�C�Y���擾
    W_UF = Me.Width
    H_UF = Me.Height
    
    '**UF�̕\���ʒu���v�Z
    T_UF = T_AW + ((H_AW - H_UF) / 2)
    L_UF = L_AW + ((W_AW - W_UF) / 2)
    
    '**UF�̕\���ʒu��ݒ�
    Me.StartUpPosition = 0
    '**Top,Left�w�莞�ɕK�{(�Ȃ���Left�������)
    Me.Top = T_UF
    Me.Left = L_UF
End Sub

