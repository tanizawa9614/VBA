Attribute VB_Name = "F_SEQUENCE2"
Option Explicit

Function SEQUENCE2(開始 As Double, _
      終了 As Double, _
      目盛り As Variant, _
      Optional _
      列方向 As Boolean = False)
   Dim row_count As Long
   Dim col_count As Long
      
   If 目盛り = 0 Then End
   If 列方向 Then
      row_count = 1
      col_count = (終了 - 開始) / 目盛り + 1
   Else
      row_count = (終了 - 開始) / 目盛り + 1
      col_count = 1
   End If
   SEQUENCE2 = WorksheetFunction.Sequence(row_count, col_count, 開始, 目盛り)
End Function
