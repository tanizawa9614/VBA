Attribute VB_Name = "S_書式調整"
Option Explicit

Sub 書式調整()
    Dim rng As Range
    Dim i As Long
    Dim myRng As Range
    Dim myText As String
    Dim myText0 As String
    Dim tmp As String, tmp2 As String
    Dim After_String
    Dim Before_String
    Dim strlen As Long

    
    After_String = Array("，", "．")
    Before_String = Array("、", "。")
    
    Set rng = Selection
    
    For Each myRng In rng
        myText = myRng.Value
        myText0 = myRng.Value
        strlen = Len(myText)
        If myText = "" Then GoTo L1
        
        For i = 0 To 9
            myText = Replace(myText, StrConv(i, vbWide), i)
        Next
        
        For i = 0 To UBound(After_String)
            myText = Replace(myText, Before_String(i), After_String(i))
        Next
        
        tmp2 = ""
        For i = 1 To strlen
            tmp = Mid(myText, i, 1)
            If tmp Like "[ａ-ｚ]" Or tmp Like "[Ａ-Ｚ]" Then
                tmp = StrConv(tmp, vbNarrow)
            End If
            tmp2 = tmp2 & tmp
        Next
        myText = tmp2
        
        For i = 1 To strlen
            If Mid(myText, i, 1) <> Mid(myText0, i, 1) Then
'                MsgBox myRng.Characters(i, 1).Text & Mid(myText, i, 1)
                myRng.Characters(i, 1).Text = Mid(myText, i, 1)
            End If
        Next
L1:
    Next
End Sub

