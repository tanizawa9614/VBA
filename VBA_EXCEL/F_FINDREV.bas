Attribute VB_Name = "F_FINDREV"
Option Explicit

Function FINDREV(‘ÎÛ, ŒŸõ•¶Žš—ñ, Optional ŠJŽnˆÊ’u As Long = 1)
    Dim findarr()
    Dim obj()
    Dim i As Long
    Dim j As Long
    Dim ans()
    Dim r As Long
    Dim c As Long
    Dim findstring
    Dim maxidx As Long
    
    Call TakeCellValue(ŒŸõ•¶Žš—ñ, findarr)
    Call TakeCellValue(‘ÎÛ, obj)
    
    r = UBound(obj, 1)
    c = UBound(obj, 2)
    ReDim ans(1 To r, 1 To c)
    
    For i = 1 To r
        For j = 1 To c
            maxidx = 0
            For Each findstring In findarr
                maxidx = Max(maxidx, InStrRev(obj(i, j), findstring))
            Next
            ans(i, j) = maxidx
        Next
    Next
    
    FINDREV = ans
End Function

Private Function Max(a, b)
    If a > b Then
        Max = a
    Else
        Max = b
    End If
    
End Function



Private Sub TakeCellValue(var, ByRef ans)
    Dim tmp
    On Error Resume Next
    tmp = var.Value2
    If Err.Number > 0 Then tmp = var
    If IsArray(tmp) = False Then
        ReDim ans(1 To 1, 1 To 1)
        ans(1, 1) = tmp
    Else
        ans = tmp
    End If
    On Error GoTo 0
End Sub
