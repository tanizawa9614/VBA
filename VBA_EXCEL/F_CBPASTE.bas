Attribute VB_Name = "F_CBPASTE"
Option Explicit

Function CBPASTE(Optional �J�n�ԍ� As Long, Optional �I���ԍ� As Long)
   Dim CB As New DataObject, i As Long
   Dim A(), j As Long
   
   CB = Application.ClipboardFormats
   ReDim A(UBound(CB))
   If CB(1) = True Then
      MsgBox "�N���b�v�{�[�h�͋�ł��B", 48
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
