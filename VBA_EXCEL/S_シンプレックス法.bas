Attribute VB_Name = "S_シンプレックス法"
Option Explicit

Sub シンプレックス法()
   Dim x_n As Long, a_n As Long
   Dim A As Range, R1, C1, B, E, z
   Dim row_sum As Long, col_sum As Long
   x_n = 3 '変数の数
   a_n = 3 '制約式の数
   row_sum = a_n + 1
   col_sum = a_n + x_n + 2
   Set A = Range("B2").CurrentRegion
   R1 = A.Resize(1)
   C1 = A.Resize(, 1)
   B = A.Resize(row_sum, col_sum).Offset(1, 1)
   z = A.Resize(1, col_sum).Offset(a_n + 1)
   For i = 1 To col_sum
      If z(i) = Min(z) Then
         
      End If
   Next
End Sub
Function 増加限界の算出()
   
End Function

Function 単位行列の作成(i As Long)
   Dim E1(), i1 As Long, j1 As Long
   ReDim E1(i - 1, i - 1)
   For i1 = 0 To i - 1
      For j1 = 0 To i - 1
         If i1 = j1 Then
            E1(i1, j1) = 1
         Else
            E1(i1, j1) = 0
         End If
      Next
   Next
   単位行列の作成 = E1
End Function
