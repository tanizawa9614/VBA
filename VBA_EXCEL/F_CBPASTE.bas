Attribute VB_Name = "F_CBPASTE"
Option Explicit

Function CBPASTE(Optional 開始番号 As Long, Optional 終了番号 As Long)
   Dim CB As New DataObject, i As Long
   Dim A(), j As Long
   
   CB = Application.ClipboardFormats
   ReDim A(UBound(CB))
   If CB(1) = True Then
      MsgBox "クリップボードは空です。", 48
      Exit Function
   End If
   For i = 1 To UBound(CB)
      If CB(i) = xlClipboardFormatBitmap Then
         ActiveSheet.Paste
      Else
         A(j) = CB(i)
         j = j + 1
      End If
   Next i
   ReDim Preserve A(j - 1)
   CBPASTE = WorksheetFunction.Transpose(A)
End Function
