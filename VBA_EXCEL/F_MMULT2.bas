Attribute VB_Name = "F_MMULT2"
Option Explicit
Option Base 1
Function MMULT2(ParamArray A())
   Dim B, i As Long, n As Long
   With WorksheetFunction
      B = A(0)
      For i = 0 To UBound(A)
         If i <> UBound(A) Then _
               B = .MMult(B, A(i + 1))
      Next
   End With
   MMULT2 = B
End Function
Function MEXPONENTIAL(ëŒè€çsóÒ, èÊêî As Long)
   Dim A, i As Long, n As Long
   A = ëŒè€çsóÒ
   n = èÊêî
   If n = 1 Then
      MEXPONENTIAL = A
      Exit Function
   End If
   For i = 2 To n
      A = WorksheetFunction.MMult(A, ëŒè€çsóÒ)
   Next
   MEXPONENTIAL = A
End Function
Function MDIAGONALIZATION(A, P)
   Dim B, i As Long, n As Long
   With WorksheetFunction
      B = .MInverse(P)
      B = .MMult(B, A)
      B = .MMult(B, P)
   End With
   MDIAGONALIZATION = B
End Function
Function IDENTITY_MATRIX(n As Long)
   Dim A(), i As Long
   ReDim A(n, n)
   For i = 1 To n
      A(i, i) = 1
   Next
   IDENTITY_MATRIX = A
End Function
