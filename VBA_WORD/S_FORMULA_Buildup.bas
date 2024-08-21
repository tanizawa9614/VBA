Attribute VB_Name = "S_FORMULA_Buildup"
Option Explicit
Sub FORMULA_Buildup()

    Dim InitRange As Word.Range
    Dim objRange As Word.Range
    Dim objOMFun As Word.OMathFunction
    Dim objSEL As Word.Selection
    
    Set InitRange = Selection.Range
    Set objRange = Selection.OMaths.Add(InitRange)

'    Set objOMFun = objRange.OMaths(1).Functions.Add(objRange, wdOMathFunctionMat)
    Set objSEL = Selection  '�����ʒu�擾
    
    '1�����E�ֈړ����Đ����Z�b�g
'    objSEL.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="(\sigma _x+\delta \sigma _x)"
    Set objSEL = Selection  '�����ʒu�擾
    
    '�����̉�ʕҏW
    objRange.OMaths(1).BuildUp
    
End Sub
Sub Test_Sample_Miniature()

    Dim InitRange As Word.Range
    Dim objRange As Word.Range
    Dim objOMFun As Word.OMathFunction
    Dim objSEL As Word.Selection
    
    '�����֐����I�u�W�F�N�g��ݒ�
    ActiveDocument.Bookmarks("\EndOfDoc").Select
    Selection.TypeParagraph
    Set InitRange = Selection.Range
    InitRange.Text = " "
    Set objRange = Selection.OMaths.Add(InitRange)

    '������ǉ�
    Set objOMFun = objRange.OMaths(1).Functions.Add(objRange, wdOMathFunctionNary)
    Set objSEL = Selection  '�����ʒu�擾
    
    objOMFun.Nary.Char = 8721
    objOMFun.Nary.HideSub = True
    objOMFun.Nary.HideSup = True
    
    '1�����E�ֈړ����Đ����Z�b�g
    objSEL.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="a/b"
    Set objSEL = Selection  '�����ʒu�擾
    
    '1�����E�ֈړ����Đ����Z�b�g
    objSEL.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="=x"
    Set objSEL = Selection  '�����ʒu�擾
    
    '�����̉�ʕҏW
    objRange.OMaths(1).BuildUp
    
    '�㎮��ǉ�
    Set objOMFun = Selection.OMaths(1).Functions.Add(Selection.Range, wdOMathFunctionRad)
    objOMFun.Rad.HideDeg = True
    Set objSEL = Selection  '�����ʒu�擾
    
    '1�������ֈړ����ă��[�g���֐����Z�b�g
    objSEL.MoveLeft Unit:=wdCharacter, Count:=1
    objSEL.TypeText Text:="x+1"
    Set objSEL = Selection  '�����ʒu�擾
    
    '1�������ֈړ����ĕ��ꐔ���Z�b�g
    objSEL.MoveRight Unit:=wdCharacter, Count:=1
    objSEL.TypeText Text:="/(a+b)"
    
    '�����̉�ʕҏW
    objRange.OMaths(1).BuildUp
    
    '�ړ�100�����ŗ��O�ւʂ���B
    On Error Resume Next
    Selection.MoveRight Unit:=wdCharacter, Count:=100
    
End Sub

