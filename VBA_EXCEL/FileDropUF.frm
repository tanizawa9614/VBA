VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileDropUF 
   Caption         =   "�t�@�C���������Ƀh���b�v���Ă�������"
   ClientHeight    =   700
   ClientLeft      =   -140
   ClientTop       =   -600
   ClientWidth     =   480
   OleObjectBlob   =   "FileDropUF.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "FileDropUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    With ListView1
    
        ''''�v���p�e�B�̐ݒ�
        .FullRowSelect = True           '�s�S�̂̑I��
'        .Gridlines = True               '�s��O���b�h���̕\��
'        .View = lvwReport               '�\���`��
        .OLEDropMode = ccOLEDropManual  '�t�@�C���h���b�v����

        ''''�񌩏o���̖��O�E�񕝂̐ݒ�
'        .ColumnHeaders.Add , "key1", "�����Ƀt�@�C�����h���b�v���Ă�������", 450, lvwColumnLeft
        .Width = 1000#
        .Height = 1000
        .Left = -10
        .BackColor = RGB(150, 150, 150)
    End With
    Me.Caption = "�t�@�C�����h���b�v����ƎQ�l�����p�̃e�L�X�g���o�͂���܂�"
    Me.Width = 400
    Me.Height = 300
End Sub


Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim i As Long
    Dim FileCount As Long
    Dim Arr() As String
    
    
    With ListView1
        
        '�t�@�C�����̎擾�i�����t�@�C���𓯎��Ƀh���b�O���h���b�v�������p�j
        FileCount = Data.Files.Count
        ReDim Arr(1 To FileCount, 1 To 1)
                
'        �h���b�O& �h���b�v�����t�@�C���p�X�����Ƀ��X�g��
        For i = 1 To FileCount
            Arr(i, 1) = Data.Files(i)
'            .ListItems.Add = Data.Files(i)
        Next i
    End With
    Unload Me
    �e�L�X�g�t�@�C���ǂݍ���_Uf (Arr)
End Sub
