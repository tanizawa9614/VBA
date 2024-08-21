Attribute VB_Name = "S_数式表示2"
Option Base 1
Sub 数式削除()
    Dim s As Range
    On Error Resume Next
    For Each s In Selection
        If Not s.Comment Is Nothing Then
            s.Comment.Delete
        End If
    Next
End Sub
Sub 数式表示()
    Dim targetC As Range
    Application.ScreenUpdating = False
    On Error Resume Next
    For Each targetC In Selection
        Call 数式表示_loop(targetC)
    Next
    Application.ScreenUpdating = True
End Sub
Private Sub 数式表示_loop(r As Range)
    Dim i As Long, j As Long
    Dim start_str As Long, buf, str2 As String
    
    If Not r.Comment Is Nothing Then
        r.Comment.Delete
    End If
    r.AddComment
    With r.Comment
        .Visible = True
        .Text r.Formula2
        With .Shape.TextFrame.Characters.Font
'            .Color = RGB(0, 0, 0)
            .Size = 14
'            .Size = WorksheetFunction.Min _
'            (WorksheetFunction.Max _
            (-0.17 * ActiveWindow.Zoom + 47, 16), 46)
        End With
        With .Shape.TextFrame
            .AutoSize = msoAutoSizeShapeToFitText
        End With
'特定文字に色付け
'特定文字とは「数字」「セル範囲」「特殊記号」「関数」を表す
        For i = 1 To Len(r.Formula2)
            buf = Mid(r.Formula2, i, 1)
            With .Shape.TextFrame
                Select Case buf
                    Case "=", "+", "-", "*", "/", """" 'Black
                        .Characters(i, 1).Font.Color = RGB(0, 0, 0)
                    Case "(", ")", ",", "{", "}"       'Red
                        .Characters(i, 1).Font.Color = RGB(200, 10, 0)
                    Case "<", ">", "\", "&", "[", "]"
                        .Characters(i, 1).Font.Color = RGB(102, 0, 204)
                    Case Else
                        If Not buf Like "#" Then
                            str2 = NextString(Mid(r.Formula2, i))
                            If IsAddress(str2) = True Then 'アドレスの場合
                                .Characters(i, Len(str2)).Font.Color = RGB(0, 160, 0)
                            ElseIf Mid(r.Formula2, i + Len(str2), 1) = "(" Then '関数の場合
                                .Characters(i, Len(str2)).Font.Color = RGB(0, 0, 220)
                            End If
                            i = i + Len(str2) - 1
                        End If
                End Select
            End With
        Next
    End With
End Sub
Private Function IsAddress(s As String) As Boolean
    On Error Resume Next
    Dim buf
    buf = Range(s).Address(False, False)
    On Error GoTo 0
    If buf <> "" Or s = "TRUE" Or s = "FALSE" Then
        IsAddress = True
        Exit Function
    End If
    IsAddress = False
End Function
Private Function NextString(str As String) As String
    Dim CSt, i As Long, j As Long, buf, mystr As String
    Dim cnt As Long
    CSt = Array(",", "(", ")", "+", "-", "*", "/", "=", ">", "<", "?", """", "&")
    mystr = Left(str, 1)
    For i = 2 To Len(str)
        buf = Mid(str, i, 1)
        For j = LBound(CSt) To UBound(CSt)
            If buf = CSt(j) Then
                Exit For
            End If
        Next
        If j = UBound(CSt) + 1 Then
            mystr = mystr & buf
        Else
            Exit For
        End If
    Next
    NextString = mystr
End Function

