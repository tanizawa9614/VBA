Attribute VB_Name = "F_FINDRIGHT"
Option Explicit

Function FINDRIGHT(‘ÎÛ As Range, ŒŸõ•¶š As Variant)
   Dim i As Long, A As Object, myA(), j As Long
   ReDim myA(‘ÎÛ.Count - 1)
   On Error Resume Next
   ŒŸõ•¶š = Left(ŒŸõ•¶š, Len(ŒŸõ•¶š))
   For Each A In ‘ÎÛ
      For i = Len(A) To 1 Step -1
         If Mid(A, i, 1) = ŒŸõ•¶š Then
            myA(j) = i
            Exit For
         End If
      Next i
      j = j + 1
   Next A
   FINDRIGHT = WorksheetFunction.Transpose(myA)
End Function
