Attribute VB_Name = "S_�����p�e�[�u���̍쐬"
Option Explicit

Sub �����p�e�[�u���̍쐬()
    
    Dim tbl As Table
    Dim bord As border
    Dim cl As Cell

    ' �i���̃X�^�C����W���ɐݒ�
    Selection.ParagraphFormat.Style = ActiveDocument.Styles("�W��")

    ' �V�����e�[�u����}��
    Set tbl = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=1, NumColumns:=3, _
                     DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed)
    
    ' �e�[�u���X�^�C����ݒ�
    tbl.Style = "�\ (�i�q)"
    
    With tbl.Range
        .Font.Name = "�l�r ����"
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        ' �Z���̔z�u�ƌr����ݒ�
        For Each cl In .Cells
            cl.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            cl.VerticalAlignment = wdCellAlignVerticalCenter
            cl.Height = 30
            For Each bord In cl.Borders
                bord.LineStyle = wdLineStyleNone
            Next
        Next
    End With
    
    tbl.Cell(2, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
        
End Sub


