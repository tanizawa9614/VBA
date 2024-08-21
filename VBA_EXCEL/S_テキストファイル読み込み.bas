Attribute VB_Name = "S_�e�L�X�g�t�@�C���ǂݍ���"
Option Explicit
Dim Author() As String
Dim Title() As String
Dim Journal() As String
Dim Vol() As Long
Dim PubPage() As Long
Dim PubYear() As Long
Dim LfFlg As Boolean
Dim LfArr
Dim FirstStr As String '�e�t�@�C���̈�s��

Sub �e�L�X�g�t�@�C���ǂݍ���()
    Dim FilePath As String
    Dim FileContent As String
    Dim LineArray() As String
    Dim Line As String
    Dim FileNum As Long
    Dim Arr
    Dim i As Long
    Dim fType As Long
    Dim n As Long
    
    Arr = Sheets(7).Cells(1, 1).CurrentRegion
    n = UBound(Arr, 1)
    
    ReDim Author(1 To n, 1 To 10)
    ReDim Title(1 To n, 1 To 1)
    ReDim Journal(1 To n, 1 To 1)
    ReDim Vol(1 To n, 1 To 2)
    ReDim PubPage(1 To n, 1 To 2)
    ReDim PubYear(1 To n, 1 To 1)
    
    For i = 1 To UBound(Arr, 1)
        ' �t�@�C���̃p�X��ݒ�
        FilePath = Arr(i, 1)
                
        ' �t�@�C�����e�L�X�g���[�h�ŊJ��
        FileNum = FreeFile
        Open FilePath For Input As FileNum
        fType = GetType(FileNum)
        Select Case fType
            Case 1
                Type1Info FileNum, i
            Case 2
                Type2Info FileNum, i
            Case 3
                MsgBox "FilePath:" & FilePath & vbCrLf & "���̃e�L�X�g�t�@�C���̓T�|�[�g����Ă��܂���" & vbCrLf & "RIS�`����������BibTeX�`����Citation����肵�Ă�������", vbCritical
        End Select
        
        ' �t�@�C�������
        Close FileNum
    Next
    Sheets(1).Cells(1, 1).Resize(n, 10) = Author
    Sheets(2).Cells(1, 1).Resize(n, 1) = Title
    Sheets(3).Cells(1, 1).Resize(n, 1) = Journal
    Sheets(4).Cells(1, 1).Resize(n, 2) = Vol
    Sheets(5).Cells(1, 1).Resize(n, 2) = PubPage
    Sheets(6).Cells(1, 1).Resize(n, 1) = PubYear
End Sub

Private Sub Type1Info(ByVal FileNum As Long, ByVal i As Long)
    Dim myLine As String
    Dim str As String
    Dim j As Long
    Dim nAuthor As Long
    
    T1Arr�o�^ FirstStr, i, nAuthor '�󔒂łȂ���s�ڂ���������
    
    If LfFlg Then
        For j = 0 To UBound(LfArr)
            myLine = LfArr(j)
            If myLine <> "" Then T1Arr�o�^ myLine, i, nAuthor
        Next
    Else
        Do Until EOF(FileNum)
            Line Input #FileNum, myLine ' 1�s�ǂݍ��� �f
            If myLine <> "" Then T1Arr�o�^ myLine, i, nAuthor
        Loop
    End If
End Sub

Private Sub T1Arr�o�^(ByVal str As String, ByVal i As Long, ByRef nAuthor As Long)
    Dim header As String, tmp
    
    header = Left(str, 2)
    tmp = Mid(str, InStr(str, "-") + 1)
    str = SpaceDeleter(tmp)
    
    Select Case header
        Case "AU"
            nAuthor = nAuthor + 1
            Author(i, nAuthor) = str
        Case "PY"
            PubYear(i, 1) = Val(Split(str, "/")(0))
        Case "TI", "T1"
            Title(i, 1) = str
        Case "JO"
            Journal(i, 1) = str
        Case "SP"
            PubPage(i, 1) = Val(str)
        Case "EP"
            PubPage(i, 2) = Val(str)
        Case "VL"
            Vol(i, 1) = Val(str)
        Case "IS"
            Vol(i, 2) = Val(str)
    End Select
End Sub

Private Sub Type2Info(ByVal FileNum As Long, ByVal i As Long)
    Dim myLine As String
    Dim str As String
    Dim j As Long
    Dim nAuthor As Long
    
    T2Arr�o�^ FirstStr, i, nAuthor '�󔒂łȂ���s�ڂ���������
    
    If LfFlg Then
        For j = 0 To UBound(LfArr)
            myLine = LfArr(j)
            If myLine <> "" Then T2Arr�o�^ myLine, i, nAuthor
        Next
    Else
        Do Until EOF(FileNum)
            Line Input #FileNum, myLine ' 1�s�ǂݍ��� �f
            If myLine <> "" Then T2Arr�o�^ myLine, i, nAuthor
        Loop
    End If
End Sub

Private Sub T2Arr�o�^(ByVal str As String, ByVal i As Long, ByRef nAuthor As Long)
    Dim header As String, tmp
    Dim j As Long
    
    tmp = Split(str, "{")
    header = SpaceDeleter(tmp(0))
    If UBound(tmp) = 0 Then Exit Sub
    tmp = Split(tmp(1), "}")
    str = SpaceDeleter(tmp(0))
    If str = "" Then Exit Sub
   
    If header Like "author*" Then
        tmp = Split(str, " and ")
        For j = 0 To UBound(tmp)
            str = SpaceDeleter(tmp(j))
            Author(i, j + 1) = str
        Next
    ElseIf header Like "year*" Then
        PubYear(i, 1) = Val(Split(str, "/")(0))
    ElseIf header Like "title*" Then
        Title(i, 1) = str
    ElseIf header Like "journal*" Then
        Journal(i, 1) = str
    ElseIf header Like "pages*" Then
        tmp = Split(str, "-")
        If UBound(tmp) < 1 Then Exit Sub
        PubPage(i, 1) = Val(SpaceDeleter(tmp(0)))
        PubPage(i, 2) = Val(SpaceDeleter(tmp(1)))
    ElseIf header Like "volume*" Then
        Vol(i, 1) = Val(str)
    ElseIf header Like "issue*" Then
        Vol(i, 2) = Val(str)
    End If
End Sub

Private Function GetType(ByVal FileNum As Long) As Long
    Dim myLine As String
    Dim tmp, LineCnt As Long
    
    Do
        Line Input #FileNum, myLine  ' �s��ǂݍ���
    Loop While myLine = ""
    
    tmp = SplitLF(myLine)
    tmp = Split(tmp(0), "-")
    tmp = SpaceDeleter(tmp(0))
    If tmp = "TY" Then
        GetType = 1  ' Type1��TY����n�܂�
    ElseIf tmp Like "@article{*" Then
        GetType = 2   ' bib�`���͈�s�ڂ�@article{��������Ă���
    Else
        GetType = 3  ' ��ԓ������Ȃ�
    End If
    
End Function

Private Function SplitLF(ByVal str)
    Dim tmp
    tmp = Split(str, vbLf)
    FirstStr = str
    If UBound(tmp) < 1 Then
        LfFlg = False
    Else
        LfFlg = True
        LfArr = tmp
    End If
    SplitLF = tmp
End Function

Private Function SpaceDeleter(ByVal str As String) As String
    Do While Left(str, 1) = " "
        str = Mid(str, 2)
    Loop
    Do While Right(str, 1) = " "
        str = Left(str, Len(str) - 1)
    Loop
    SpaceDeleter = str
End Function
