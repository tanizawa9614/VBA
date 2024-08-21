Attribute VB_Name = "F_ConvertToZeroBasedArray"
Option Explicit
'Sub test()
'    Dim a(0 To 3), b(1 To 3, 1 To 3), c(1 To 3)
'    Dim a1, a2, a3
'    a1 = ConvertToZeroBasedArray(a)
'    a2 = ConvertToZeroBasedArray(b)
'    a3 = ConvertToZeroBasedArray(c)
'End Sub

Function ConvertToZeroBasedArray(ByVal matA As Variant) As Variant
    Dim rows As Integer
    Dim cols As Integer
    Dim zeroBasedArray() As Double
    Dim ansA
    Dim i As Integer, j As Integer
    
    On Error Resume Next
        cols = UBound(matA, 2) - LBound(matA, 2) + 1
    If Err.Number = 0 Then
        rows = UBound(matA, 1) - LBound(matA, 1) + 1
        
        ReDim zeroBasedArray(0 To rows - 1, 0 To cols - 1)
    
        For i = LBound(matA, 1) To UBound(matA, 1)
            For j = LBound(matA, 2) To UBound(matA, 2)
                zeroBasedArray(i, j) = matA(i, j)
            Next j
        Next i
        ansA = zeroBasedArray
    Else
        rows = 1
        cols = UBound(matA) - LBound(matA) + 1
        Dim tempArr() As Double
        ReDim tempArr(0 To cols - 1, 0)
        For i = LBound(matA) To UBound(matA)
            tempArr(i - LBound(matA), 0) = matA(i)
        Next i
        ansA = tempArr
    End If
    On Error GoTo 0
    
    
    
    ConvertToZeroBasedArray = ansA
    
End Function

