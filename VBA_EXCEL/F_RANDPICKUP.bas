Attribute VB_Name = "F_RANDPICKUP"
Option Explicit

Function RANDPICKUP(配列, 行数 As Long)
    Dim arr, num_rows As Long
    Dim id_arr() As Long, ans
    Dim cnt As Long, i As Long
    Dim id As Long

    Randomize
    arr = 配列
    num_rows = 行数
    ReDim id_arr(1 To num_rows)
    ReDim ans(1 To num_rows, 1 To UBound(arr, 2))
    If num_rows >= UBound(arr, 1) Then
        MsgBox "入力の配列のサイズ（行数）より小さい値を第二引数に渡してください"
        Exit Function
    End If
    cnt = 1
    Do While cnt <= num_rows
        id = Int((UBound(arr, 1) - 1 + 1) * Rnd + 1)
        
        ' すでにリストにないかチェック
        If IsInArray(id, id_arr) = False Then
            For i = 1 To UBound(arr, 2)
                ans(cnt, i) = arr(id, i)
                id_arr(cnt) = id
            Next
            cnt = cnt + 1
        End If
    Loop
    
    RANDPICKUP = ans
End Function

' 配列内に特定の値が存在するかどうかをチェックする関数
Private Function IsInArray(val As Long, arr() As Long) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

