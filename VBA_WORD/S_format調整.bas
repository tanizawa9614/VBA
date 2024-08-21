Attribute VB_Name = "S_format調整"
Option Explicit
Dim start_range As Long, end_range As Long
Sub Replace_to_da()
    Dim A
    Dim B, i As Long
    A = Array("でした", "ります", "います", "します", "なりました", "行いました", "わかりました", "分かりました", "ました", "ありませんでした", "ありません", "です", "ます", "ません")
    B = Array("であった", "る", "いる", "する", "なった", "行った", "分かった", "分かった", "た", "なかった", "ない", "である", "る", "ない")
    
    For i = 0 To UBound(A)
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = A(i)
            .Replacement.Text = B(i)
            .Forward = True
            .Wrap = wdFindContinue
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next
End Sub

Sub format調整_main()
    Dim s As Object, i As Long, j As Long
    Dim start_n As Long, end_n As Long
    
     
    Set s = Selection
    start_n = s.Range.Start
    end_n = s.Range.End
    
    DoEvents
    For i = start_n To end_n
        If Cer_Font(i) = False Then
            start_range = i
            For j = i To end_n
                If Cer_Font(j) = True Then
                    end_range = j
                    Exit For
                End If
                If j = end_n Then end_range = j
            Next
            Call format調整
            i = end_range + 1
        End If
    Next
    ActiveDocument.Range(start_n, end_n).Select
End Sub
Private Function Cer_Font(n As Long)
    
    With ActiveDocument.Range(n, n + 1).Font
        Select Case True
        Case .Italic, .Bold, .Underline, .Subscript, .Superscript
            Cer_Font = True
            Exit Function
        End Select
    End With
    If n + 1 = ActiveDocument.Range.End Then
        end_range = n
        Call format調整
        End
    End If
    Cer_Font = False
End Function

Private Sub format調整()
    Dim n As Long, i As Long
    Dim txt As String
    Dim str As String
    Dim tmp As String
    
    str = ActiveDocument.Range(start_range, end_range).Text
    
    If str = vbCr Or str = vbLf Or str = vbCrLf Then Exit Sub
    
    n = Len(str)
    str = Replace(str, "、", "，")
    str = Replace(str, "。", "．")
    str = Replace(str, "　", "")
    
    For i = 1 To n
        txt = Mid(str, i, 1)
        
        Select Case txt
        Case " "
            tmp = ActiveDocument.Range(start_range + i - 2, start_range + i - 1)
            If tmp <> "," And tmp <> "." Then
                If 前後が半角(str, i) = False Then txt = ""
            End If
        Case ","
            If 前後が半角(str, i) = False Then txt = "，"
        Case "(", ")"
            If 前後が半角(str, i) = False Then txt = StrConv(txt, vbWide)
        Case "（", "）"
            If 前後が半角(str, i) = True Then txt = StrConv(txt, vbNarrow)
        Case "，"
            If 前後が半角(str, i) = True Then txt = ","
        Case "."
            If 前後が半角(str, i) = False Then txt = "．"
        Case "．"
            If 前が英文字(str, i) = True Then txt = "."
        End Select
        
        str = Left(str, i - 1) & txt & Mid(str, i + 1)
        n = Len(str)
        If i = n Then Exit For
        i = i - (1 - Len(txt))
    Next
    ActiveDocument.Range(start_range, end_range).Text = str
End Sub
Private Function make_str(str As String, i As Long, ope As String)
    Dim cnt As Long, tmp As String
    cnt = 1
    Do While tmp = " " Or tmp = "　" Or tmp = ""
        If ope = "+" Then
            tmp = ActiveDocument.Range(start_range + i + cnt - 1, start_range + i + cnt).Text
            If i + cnt = Len(Selection.Text) Then Exit Do
        Else
            tmp = ActiveDocument.Range(start_range + i - cnt - 1, start_range + i - cnt).Text
            If i - cnt = 1 Then Exit Do
        End If
        cnt = cnt + 1
    Loop
    make_str = tmp
End Function
Private Function 前後が半角(str As String, i As Long) As Boolean
    Dim bef As String, aft As String
    Dim flg1 As Boolean, flg2 As Boolean
    bef = make_str(str, i, "-")
    aft = make_str(str, i, "+")
    
    If str_check(bef) = False Then
        前後が半角 = False
        Exit Function
    End If
    If str_check(aft) = False Then
        前後が半角 = False
        Exit Function
    End If
    前後が半角 = True
End Function
Private Function 前が英文字(str As String, i As Long) As Boolean
    Dim bef As String
    bef = make_str(str, i, "-")
    If str_check(bef) = False Then
        前が英文字 = False
        Exit Function
    End If
    前が英文字 = True
End Function
Private Function str_check(str As String) As Boolean
    Dim A, i As Long, buf As String
    A = Array("(", ")", "+", "*", "-", _
                "^", "/", "<", ">", "&", _
                "#", "%", "@", ";", ":", "{", "}")
    
    If str Like "[A-Z]" Or str Like "[a-z]" Then
        str_check = True
        Exit Function
    End If
    
    For i = 0 To 9
        buf = "" & i
        If str = buf Then
            str_check = True
            Exit Function
        End If
    Next
    
    For i = 0 To UBound(A)
        If str = A(i) Then
            str_check = True
            Exit Function
        End If
    Next
    str_check = False
End Function

