Attribute VB_Name = "S_SUBSTITUTE2"
Option Explicit
Function SUBSTITUTE2_ファイル名体裁用(文字列)

    Dim strarr
    strarr = Array(",", "\", "/", ":", ";", "*", "?", "<", ">", "|", " ", "-", vbTab, vbLf, vbCr)
    SUBSTITUTE2_ファイル名体裁用 = SUBSTITUTE2(文字列, strarr, " ")
    
End Function

Function SUBSTITUTE2(文字列, 検索文字列, 置換文字列)
    Dim arr, findA, repA
    Dim i As Long
    Dim j As Long
    Dim tmpstr As String

    ' 文字列が配列の場合
    If IsArray(文字列) Then
        arr = 文字列
        ' 配列の各要素に対して処理
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                tmpstr = arr(i, j)
                arr(i, j) = ReplaceMultiple(tmpstr, 検索文字列, 置換文字列)
            Next
        Next
    Else ' 文字列が変数の場合
        tmpstr = 文字列
        arr = ReplaceMultiple(tmpstr, 検索文字列, 置換文字列)
    End If
    
    ' 置換後の結果を出力する (例としてセルA1に結果を表示)
    SUBSTITUTE2 = arr
End Function

Private Function ReplaceMultiple(inputStr As String, findArr, repStr) As String
    Dim i As Long
    Dim findStr As Variant

    ' 検索文字列が配列の場合
    If IsArray(findArr) Then
        For Each findStr In findArr
            inputStr = Replace(inputStr, findStr, repStr)
        Next findStr
    Else ' 検索文字列が変数の場合
        inputStr = Replace(inputStr, findArr, repStr)
    End If

    ReplaceMultiple = inputStr
End Function

