Attribute VB_Name = "S_テキストファイル読み込み2"
Option Explicit
Dim Author() As String
Dim Title() As String
Dim Journal() As String
Dim Vol() As Long
Dim PubPage() As Long
Dim PubYear() As Long
Dim LfFlg As Boolean
Dim LfArr
Dim FirstStr As String '各ファイルの一行目
Dim adoSt As Object

Sub テキストファイル読み込み_Uf(Arr)
    Dim FilePath As String
    Dim FileContent As String
    Dim LineArray() As String
    Dim Line As String
    Dim FileNum As Long
    Dim i As Long
    Dim fType As String
    Dim n As Long
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim sh As Worksheet
    Application.ScreenUpdating = False
    With ActiveWorkbook
        Set sh = .Sheets.Add(After:=.Sheets(1))
    End With
    Dim tmp As String
    tmp = Format(Now, "yyyy-mm-dd hh時mm分ss秒")
    sh.Name = tmp
    n = UBound(Arr, 1)
    
    ReDim Author(1 To n, 1 To 10)
    ReDim Title(1 To n, 1 To 1)
    ReDim Journal(1 To n, 1 To 1)
    ReDim Vol(1 To n, 1 To 2)
    ReDim PubPage(1 To n, 1 To 2)
    ReDim PubYear(1 To n, 1 To 1)
    
    'ADODB.Streamオブジェクトを生成
    Set adoSt = CreateObject("ADODB.Stream") ' 文字化けを阻止するために通常の読み込みは使わない
    adoSt.Charset = "UTF-8"
    adoSt.Open
    
    For i = 1 To UBound(Arr, 1)
        ' ファイルのパスを設定
        FilePath = Arr(i, 1)
        If Dir(FilePath) = "" Then GoTo L1
        If FilePath = "" Then GoTo L1
        
        fType = FSO.GetExtensionName(FilePath)
        If fType <> "ris" And fType <> "bib" Then
            MsgBox "FilePath:" & vbCrLf & FilePath & vbCrLf & "このファイルはサポートされていません" & vbCrLf & "RIS形式もしくはBibTeX形式のCitationを入手してください", vbCritical
            GoTo L1
        End If
        adoSt.LoadFromFile (FilePath)
                
        Call GetType
        
        Select Case fType
            Case "ris" '1
                Type1Info FileNum, i   ' RIS形式
            Case "bib" '2
                Type2Info FileNum, i   ' BibTeX形式
        End Select
L1:
    Next
    
    adoSt.Close
    
    Dim ansStr()
    ReDim ansStr(1 To n, 1 To 1)
    Dim AuthorStr As String
    Dim VolStr As String
    Dim PageStr As String
    Dim YearStr As String
    Dim ext As String
    
    Dim NewFileName As String
    
    For i = 1 To n
        AuthorStr = MakeAuthorStr(i)
        VolStr = MakeVolStr(i)
        PageStr = MakePageStr(i)
        YearStr = MakeYearStr(i)
        ansStr(i, 1) = AuthorStr & Title(i, 1) & ". " & Journal(i, 1) & "." & VolStr & PageStr & YearStr
        ansStr(i, 1) = Replace(ansStr(i, 1), ",,", ",")
        ansStr(i, 1) = Replace(ansStr(i, 1), "  ", " ")
        ansStr(i, 1) = Replace(ansStr(i, 1), ",.", ",")
'        ansStr(i,1) = Replace(ansStr(i,1), ".,", ",")
        NewFileName = FSO.GetParentFolderName(Arr(i, 1)) & "\" & ansStr(i, 1)
        NewFileName = Replace(NewFileName, ";", "")
        NewFileName = Replace(NewFileName, "-", "")
        ext = FSO.GetExtensionName(Arr(i, 1))
        Do While FSO.FileExists(NewFileName & "." & ext) = True
            NewFileName = NewFileName & "_1"
        Loop
        If ext = "ris" Or ext = "bib" Then
            On Error Resume Next
            Name Arr(i, 1) As NewFileName & "." & ext
            On Error GoTo 0
        End If
    Next
    sh.Cells(1, 1).Resize(n, 1) = ansStr
    Application.ScreenUpdating = True
    Set sh = Nothing
End Sub
Private Function MakeYearStr(ByVal i As Long) As String
    Dim ans As String
    If PubYear(i, 1) <> 0 Then ans = "," & PubYear(i, 1) & "."
    If Vol(i, 1) = 0 Then ans = Mid(ans, 2)
    MakeYearStr = ans
End Function

Private Function MakeVolStr(ByVal i As Long) As String
    Dim ans As String
    If Vol(i, 1) <> 0 Then
        ans = Vol(i, 1)
        If Vol(i, 2) <> 0 Then
            ans = ans & "(" & Vol(i, 2) & "),"
        Else
            ans = ans & ","
        End If
    End If
    MakeVolStr = ans
End Function

Private Function MakePageStr(ByVal i As Long) As String
    Dim ans As String
    If PubPage(i, 2) <> 0 Then ans = "pp." & PubPage(i, 1) & "-" & PubPage(i, 2) & ","
    MakePageStr = ans
End Function
Private Function MakeAuthorStr(ByVal i As Long) As String
    Dim j As Long
    Dim ans As String
    Dim splitAuthor
    Dim tmp, tmpstr As String
    Dim k As Long
    Do While Author(i, j + 1) <> ""
        j = j + 1
        splitAuthor = Split(Author(i, j), ",")
        tmpstr = ""
        For k = 1 To UBound(splitAuthor)
            tmpstr = tmpstr & Left(splitAuthor(k), 1) & "."
        Next
        ans = ans & splitAuthor(0) & ", " & tmpstr & IIf(Author(i, j + 1) <> "", ", ", ": ")
    Loop
    MakeAuthorStr = ans
End Function


Private Sub Type1Info(ByVal FileNum As Long, ByVal i As Long)
    Dim myLine As String
    Dim str As String
    Dim j As Long
    Dim nAuthor As Long
    
    T1Arr登録 FirstStr, i, nAuthor '空白でない一行目も処理する
    
    If LfFlg Then
        For j = 0 To UBound(LfArr)
            myLine = LfArr(j)
            If myLine <> "" Then T1Arr登録 myLine, i, nAuthor
        Next
    Else
        With adoSt
            Do Until .EOF
                myLine = .Readtext() ' 行を読み込む
                If myLine <> "" Then T1Arr登録 myLine, i, nAuthor
            Loop
        End With
    End If
End Sub

Private Sub T1Arr登録(ByVal str As String, ByVal i As Long, ByRef nAuthor As Long)
    Dim header As String, tmp
    
    header = Left(str, 2)
    tmp = Mid(str, InStr(str, "-") + 1)
    str = SpaceDeleter(tmp)
    
    Select Case header
        Case "AU"
            nAuthor = nAuthor + 1
            Author(i, nAuthor) = MakeNameStr(str)
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
    
    T2Arr登録 FirstStr, i '空白でない一行目も処理する
    
    If LfFlg Then
        For j = 0 To UBound(LfArr)
            myLine = LfArr(j)
            If myLine <> "" Then T2Arr登録 myLine, i
        Next
    Else
        With adoSt
            Do Until .EOF
                myLine = .Readtext() ' 行を読み込む
                If myLine <> "" Then T2Arr登録 myLine, i
            Loop
        End With
    End If
End Sub

Private Sub T2Arr登録(ByVal str As String, ByVal i As Long)  ' BibTeX用Arr作成
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
            Author(i, j + 1) = MakeNameStr(str)
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
    ElseIf header Like "issue*" Or header Like "number*" Then
        Vol(i, 2) = Val(str)
    End If
End Sub

Private Sub GetType()
    Dim myLine As String
    Dim tmp, LineCnt As Long
    Dim LfTmp
    Dim i As Long
    
    With adoSt
        Do
            myLine = .Readtext() ' 行を読み込む
        Loop While myLine = ""
    End With
    
    LfTmp = SplitLF(myLine)
    Do
        tmp = Split(LfTmp(i), "-")
        tmp = SpaceDeleter(tmp(0))
        i = i + 1
    Loop While Len(tmp) <= 1
    
'    If tmp = "TY" Then
'        GetType = 1  ' Type1はTYから始まる,RIS形式
'    ElseIf tmp Like "@article{*" Then
'        GetType = 2   ' BibTeX形式は一行目に@article{が書かれている
'    Else
'        GetType = 3  ' 一番特徴がない
'    End If
    
End Sub

Private Function SplitLF(ByVal str)  ' Lf形式で書かれたファイルは一行で読み込まれるためLfでSplitする
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

Private Function SpaceDeleter(ByVal str As String) As String ' スペース削除を行う関数
    Do While Left(str, 1) = " "
        str = Mid(str, 2)
    Loop
    Do While Right(str, 1) = " " Or Right(str, 1) = "," Or Right(str, 1) = "."
        str = Left(str, Len(str) - 1)
    Loop
    SpaceDeleter = Replace(str, vbCrLf, "")
End Function

Private Function MakeNameStr(ByVal str As String) As String
    Dim tmp, i As Long, ans As String
    If InStr(str, ",") = 0 Then
        tmp = Split(str, " ")
        For i = 0 To UBound(tmp) - 1
            ans = ans & tmp(i) & IIf(i = UBound(tmp), "", " ")
        Next
        str = tmp(UBound(tmp)) & ", " & ans
    End If
    str = Replace(str, " ", ",")
    str = Replace(str, ",,", ",")
    MakeNameStr = SpaceDeleter(str)
End Function
