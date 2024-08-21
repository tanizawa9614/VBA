Attribute VB_Name = "F_RANDSTRING"
Option Explicit

Function RANDSTRING(çs As Long, _
   Optional óÒ As Long = 1, _
   Optional ï∂éöêî As Long = 1)
   
   Dim Unicode_Ç† As Long
   Dim Unicode_ÇÒ As Long
   Dim A(), i As Long, j As Long
   Dim k As Long
   ReDim A(çs - 1, óÒ - 1)
   
   Application.Volatile
   Unicode_Ç† = 12354
   Unicode_ÇÒ = 12435
   For i = 0 To çs - 1
      For j = 0 To óÒ - 1
         For k = 1 To ï∂éöêî
            With WorksheetFunction
               A(i, j) = A(i, j) & .Unichar(.RandBetween(Unicode_Ç†, Unicode_ÇÒ))
            End With
         Next
      Next
   Next
   RANDSTRING = A
End Function


