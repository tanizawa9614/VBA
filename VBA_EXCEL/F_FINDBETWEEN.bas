Attribute VB_Name = "F_FINDBETWEEN"
Option Explicit

Function FINDBETWEEN(対象 As Range, _
                        スタート As Variant, _
                           ストップ As Variant)
   Dim upper As Long, lower As Long
   Dim C As Range, mya(), i As Long
   ReDim mya(対象.Count - 1)
   For Each C In 対象
      upper = InStr(C, スタート)
      lower = InStr(C, ストップ)
      If upper <> 0 And lower <> 0 Then
         mya(i) = Mid(C, upper + 1, lower - upper - 1)
      Else
         mya(i) = "なし"
      End If
      i = i + 1
   Next C
   FINDBETWEEN = WorksheetFunction.Transpose(mya)
End Function
 


