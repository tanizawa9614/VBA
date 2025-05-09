Attribute VB_Name = "F_RANDSTRING"
Option Explicit

Function RANDSTRING(行 As Long, _
   Optional 列 As Long = 1, _
   Optional 文字数 As Long = 1)
   
   Dim Unicode_あ As Long
   Dim Unicode_ん As Long
   Dim A(), i As Long, j As Long
   Dim k As Long
   ReDim A(行 - 1, 列 - 1)
   
   Application.Volatile
   Unicode_あ = 12354
   Unicode_ん = 12435
   For i = 0 To 行 - 1
      For j = 0 To 列 - 1
         For k = 1 To 文字数
            With WorksheetFunction
               A(i, j) = A(i, j) & .Unichar(.RandBetween(Unicode_あ, Unicode_ん))
            End With
         Next
      Next
   Next
   RANDSTRING = A
End Function


