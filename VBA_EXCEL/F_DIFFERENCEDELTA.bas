Attribute VB_Name = "F_DIFFERENCEDELTA"
Option Explicit

Function DIFFERENCEDELTA(�z��)
    Dim arr
    arr = �z��
    Dim n_rows As Integer
    n_rows = UBound(arr, 1)
    Dim ans()
    ReDim ans(1 To n_rows - 1, 1 To 1)
    Dim i As Integer
    For i = 1 To n_rows - 1
        If IsNumeric(arr(i, 1)) And IsNumeric(arr(i + 1, 1)) Then
            ans(i, 1) = arr(i + 1, 1) - arr(i, 1)
        End If
    Next
    DIFFERENCEDELTA = ans
End Function

