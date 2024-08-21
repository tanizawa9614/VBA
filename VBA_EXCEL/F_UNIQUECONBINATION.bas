Attribute VB_Name = "F_UNIQUECONBINATION"
Option Explicit
Function UNIQUECONBINATION(配列, Optional 列数 As Long = -1)
    Dim arr, ans
    Dim num_col As Integer
    arr = 配列
    If 列数 = -1 Then
        num_col = UBound(arr, 2) - 1
    ElseIf 列数 >= UBound(arr, 2) Or 列数 < 1 Then
        UNIQUECONBINATION = arr
        Exit Function
    Else
        num_col = 列数
    End If
    Dim tmp() As Integer
    ReDim tmp(1 To num_col)
    Call GetCombinations(1, 1, UBound(arr, 2), num_col, tmp, arr, ans)
    SORTROW ans
    UNIQUECONBINATION = WorksheetFunction.Sort(WorksheetFunction.Unique(ans))
End Function

Private Sub SORTROW(ByRef arr0, Optional 降順 As Boolean = True)
    Dim ans0()
    Dim col_count As Long
    Dim i As Long
    Dim j As Long
    
    ReDim ans0(LBound(arr0) To UBound(arr0), LBound(arr0, 2) To UBound(arr0, 2))
    Dim rowarr0(), sorted_row
    ReDim rowarr0(LBound(arr0, 2) To UBound(arr0, 2))
    Dim loop_max As Integer
    loop_max = UBound(arr0, 2) - LBound(arr0, 2)
    
    ' 各行ごとにソート
    For i = LBound(arr0) To UBound(arr0)
        ' 各行の要素を配列に格納
        For j = LBound(arr0, 2) To UBound(arr0, 2)
            rowarr0(j) = arr0(i, j)
        Next
        
        If 降順 Then
            sorted_row = WorksheetFunction.Sort(rowarr0, , 1, True)
        Else
            sorted_row = WorksheetFunction.Sort(rowarr0, , -1, True)
        End If
        ' ソートされた行を出力配列に格納
        Dim cnt As Integer, id_min As Integer
        id_min = LBound(sorted_row)
        cnt = 0
        
        For j = LBound(arr0, 2) To UBound(arr0, 2)
            ans0(i, j) = sorted_row(id_min + cnt)
            cnt = cnt + 1
        Next
    Next
    arr0 = ans0
End Sub

Private Sub GetCombinations(start As Integer, index As Integer, n As Integer, M As Integer, combination() As Integer, arr, ans)
    Dim i As Integer

    If index > M Then
        ' 組み合わせが完成した場合
        Dim tmp
        tmp = ChooseCol(arr, combination)
        If IsEmpty(ans) Then
            ans = tmp
        Else
            ans = Vstack(ans, tmp)
        End If
    Else
        ' 再帰的に組み合わせを生成
        For i = start To n - M + index
            combination(index) = i
            Call GetCombinations(i + 1, index + 1, n, M, combination, arr, ans)
        Next i
    End If
End Sub

Private Function ChooseCol(ByRef arr0, ByRef col_idx() As Integer)
    Dim ans0(), i As Long, j As Long
    ReDim ans0(1 To UBound(arr0, 1), 1 To UBound(col_idx))
    For i = 1 To UBound(arr0, 1)
        For j = 1 To UBound(col_idx)
            ans0(i, j) = arr0(i, col_idx(j))
        Next
    Next
    ChooseCol = ans0
End Function

Private Function Vstack(ByRef arr0, ByRef arr1)
    Dim ans0(), i As Long, j As Long
    ReDim ans0(1 To UBound(arr0, 1) + UBound(arr1, 1), 1 To UBound(arr0, 2))
    For i = 1 To UBound(arr0, 1)
        For j = 1 To UBound(arr0, 2)
            ans0(i, j) = arr0(i, j)
        Next
    Next
    For i = 1 To UBound(arr1, 1)
        For j = 1 To UBound(arr1, 2)
            ans0(i + UBound(arr0, 1), j) = arr1(i, j)
        Next
    Next
    Vstack = ans0
End Function
