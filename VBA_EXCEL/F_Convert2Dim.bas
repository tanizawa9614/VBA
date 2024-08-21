Attribute VB_Name = "F_Convert2Dim"
Option Explicit

Public Function Convert2Dim(arr0)
    Dim ans() As Variant, arr
    Dim i As Long, j As Long
    Dim ic As Long, jc As Long
    Dim numRows As Long, numCols As Long
    
    arr = arr0
    
    If IsArray(arr) Then
        On Error Resume Next
        numRows = UBound(arr, 1) - LBound(arr, 1) + 1
        numCols = UBound(arr, 2) - LBound(arr, 2) + 1
        On Error GoTo 0
        
        If numCols = 0 Then ' 1éüå≥îzóÒÇÃèÍçá
            ReDim ans(1 To 1, 1 To numRows)
            
            For i = LBound(arr) To UBound(arr)
                ic = ic + 1
                ans(1, ic) = arr(i)
            Next i
            
        Else ' 2éüå≥îzóÒÇÃèÍçá
            If LBound(arr, 1) * LBound(arr, 2) = 1 Then
                ans = arr
            Else
                ReDim ans(1 To numRows, 1 To numCols)
                For i = LBound(arr, 1) To UBound(arr, 1)
                    ic = ic + 1
                    jc = 0
                    For j = LBound(arr, 2) To UBound(arr, 2)
                        jc = jc + 1
                        ans(ic, jc) = arr(i, j)
                    Next j
                Next i
            End If
        End If
    Else
        ' If input is not an array, create a 1x1 array with minimum index 1
        ReDim ans(1 To 1, 1 To 1)
        ans(1, 1) = arr
    End If
    
    Convert2Dim = ans
End Function
