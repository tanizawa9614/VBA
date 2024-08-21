Attribute VB_Name = "Module1"
Option Explicit

Function INPUTMESHDATA(A0)
    Dim A, i As Long, str As String
    A = A0
    A = MAKE2DARRAY(A)
    For i = 1 To UBound(A)
        str = A(i, 1)
    Next
End Function
Private Function MAKE2DARRAY(A)
    Dim i As Long, j As Long
    
End Function
