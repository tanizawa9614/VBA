Attribute VB_Name = "S_SUBSTITUTE2"
Option Explicit
Function SUBSTITUTE2_�t�@�C�����̍ٗp(������)

    Dim strarr
    strarr = Array(",", "\", "/", ":", ";", "*", "?", "<", ">", "|", " ", "-", vbTab, vbLf, vbCr)
    SUBSTITUTE2_�t�@�C�����̍ٗp = SUBSTITUTE2(������, strarr, " ")
    
End Function

Function SUBSTITUTE2(������, ����������, �u��������)
    Dim arr, findA, repA
    Dim i As Long
    Dim j As Long
    Dim tmpstr As String

    ' �����񂪔z��̏ꍇ
    If IsArray(������) Then
        arr = ������
        ' �z��̊e�v�f�ɑ΂��ď���
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                tmpstr = arr(i, j)
                arr(i, j) = ReplaceMultiple(tmpstr, ����������, �u��������)
            Next
        Next
    Else ' �����񂪕ϐ��̏ꍇ
        tmpstr = ������
        arr = ReplaceMultiple(tmpstr, ����������, �u��������)
    End If
    
    ' �u����̌��ʂ��o�͂��� (��Ƃ��ăZ��A1�Ɍ��ʂ�\��)
    SUBSTITUTE2 = arr
End Function

Private Function ReplaceMultiple(inputStr As String, findArr, repStr) As String
    Dim i As Long
    Dim findStr As Variant

    ' ���������񂪔z��̏ꍇ
    If IsArray(findArr) Then
        For Each findStr In findArr
            inputStr = Replace(inputStr, findStr, repStr)
        Next findStr
    Else ' ���������񂪕ϐ��̏ꍇ
        inputStr = Replace(inputStr, findArr, repStr)
    End If

    ReplaceMultiple = inputStr
End Function

