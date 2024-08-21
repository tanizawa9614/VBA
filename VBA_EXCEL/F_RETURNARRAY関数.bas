Attribute VB_Name = "Module1"
Option Explicit
Option Base 1
Function RETURNARRAY(範囲 As Range, 行数 As Long, _
                     Optional 代替テキスト As Variant = "")
   Dim A
   Dim i As Long, j As Long, k As Long
   On Error Resume Next
   A = 範囲
   Dim Row As Long, Col As Long, N As Long
   Row = 行数
   Col = UBound(A, 2)
   N = WorksheetFunction.RoundUp((UBound(A, 1) - 1) / (Row), 0)
   Dim B()
   ReDim B(Row + 1, Col * N + N)
   For k = 1 To N
      For i = 1 To Row + 1
         For j = 1 To Col + 1
            If j = Col + 1 Then
               B(i, j + (Col + 1) * (k - 1)) = ""
            ElseIf i = 1 Then
               B(i, j + (Col + 1) * (k - 1)) = A(1, j)
            Else
               B(i, j + (Col + 1) * (k - 1)) = A(i + Row * (k - 1), j)
               If Err.Number = 9 Then B(i, j + (Col + 1) * (k - 1)) = 代替テキスト
            End If
         Next j
      Next i
   Next k
   ReDim Preserve B(Row + 1, Col * N + N - 1)
   RETURNARRAY = B
End Function

