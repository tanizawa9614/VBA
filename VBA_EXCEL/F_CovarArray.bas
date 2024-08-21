Attribute VB_Name = "F_CovarArray"
Option Explicit

Function CovarArray(data As Variant, Optional populationVariance As Boolean = False) As Variant

    ' •ªŽUs—ñ‚ð•Ô‚µ‚½‚¢ê‡‚ÉŽg‚¤ŠÖ”
    'populationVariance‚Í•ê•ªŽU‚ð•Ô‚µ‚½‚¢‚Æ‚«‚Ì‚ÝTRUE‚ÉC•W€‚Å‚Í•s•Ï•ªŽU‚ª•Ô‚é

    Dim numRows As Long
    Dim numCols As Long
    Dim dataArray
    Dim ans() As Variant
    Dim i As Long, j As Long, k As Long
    Dim sumXY As Double, sumX As Double, sumY As Double
    Dim count As Long
    
    dataArray = data
    numRows = UBound(dataArray, 1)
    numCols = UBound(dataArray, 2)
    
    ReDim ans(1 To numCols, 1 To numCols)
    
    For i = 1 To numCols
        For j = 1 To numCols
            sumXY = 0
            sumX = 0
            sumY = 0
            count = 0
            
            For k = 1 To numRows
                If IsNumeric(dataArray(k, i)) And IsNumeric(dataArray(k, j)) Then
                    sumXY = sumXY + dataArray(k, i) * dataArray(k, j)
                    sumX = sumX + dataArray(k, i)
                    sumY = sumY + dataArray(k, j)
                    count = count + 1
                End If
            Next k
            
            If count > 1 Then
                If populationVariance Then
                    ans(i, j) = (sumXY - sumX * sumY / count) / count
                Else
                    ans(i, j) = (sumXY - sumX * sumY / count) / (count - 1)
                End If
            Else
                ans(i, j) = CVErr(xlErrNA)
            End If
        Next j
    Next i
    
    CovarArray = ans
End Function


