Attribute VB_Name = "F_GetFreeFormNodes"
Function GetFreeFormNodes(rng, xmin As Double, xmax As Double, ymin As Double, ymax As Double)
    Dim regex As Object
    Dim matches As Object
    Dim numbers() As Double
    Dim i As Long
    Dim j As Long
    Dim cell
    Dim str As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "[-+]?\d*\.?\d+"
    
    ReDim numbers(1 To 2, 1 To rng.Cells.Count) As Double
    i = 1
    
    For Each cell In rng
        str = cell.Value
        
        If InStr(1, str, "AddNode", vbTextCompare) > 0 Or InStr(1, str, "BuildFreeform", vbTextCompare) > 0 Then
            For j = 1 To 2
                Set matches = regex.Execute(str)
                If matches.Count > 0 Then
                    numbers(j, i) = CDbl(matches(0))
                End If
                On Error Resume Next
                str = Replace(str, matches(0), "")
                If Err.Number > 0 Then
                    MsgBox "数値の間に改行がある可能性があります" & vbCr & "[_]の部分を前の行に追加してやり直してください"
                    Exit Function
                End If
                On Error GoTo 0
            Next
             i = i + 1
        End If
    Next cell
    ReDim Preserve numbers(1 To 2, 1 To i - 1)
    
    Dim interpolate()
    ReDim interpolate(1 To i - 1, 1 To 2)
    Dim xdmin As Double, xdmax As Double
    Dim ydmin As Double, ydmax As Double
    Dim xd(), yd(), ans()
    ReDim xd(1 To i - 1), yd(1 To i - 1), ans(1 To i - 1, 1 To 2)
    
    
    For j = 1 To i - 1
        xd(j) = numbers(1, j)
        yd(j) = numbers(2, j)
    Next
    
    With WorksheetFunction
        xdmin = .Min(xd)
        ydmin = .Max(yd)
        xdmax = .Max(xd)
        ydmax = .Min(yd)
        For j = 1 To i - 1
            ans(j, 1) = (xmax - xmin) / (xdmax - xdmin) * (xd(j) - xdmin) + xmin
            ans(j, 2) = (ymax - ymin) / (ydmax - ydmin) * (yd(j) - ydmin) + ymin
        Next
        GetFreeFormNodes = ans
    End With
End Function

