Attribute VB_Name = "S_ä÷êîïœå`"
Option Explicit
Option Base 1
Dim Cst(4)
Sub Make_Cst()
   Cst(1) = "+"
   Cst(2) = "-"
   Cst(3) = "*"
   Cst(4) = "/"
   Cst(5) = "^"
End Sub
Sub ä÷êîïœå`()
   Const S As String = "a^2+b+c/d*a"
   Dim i As Long, j As Long, S2 As String
   Dim Split_S
   Make_Cst
   S2 = S
   For i = 1 To UBound(Cst)
      S2 = Replace(S2, Cst(i), vbTab)
   Next
   Split_S = Split(S2, vbTab)
   For i = 1 To Len(S2)
'      S2 = Left(FINDSYMBOL(Mid(S2, i, 1)))
   Next
End Sub

Function FINDSYMBOL(S As String)
   Select Case S
   Dim S2 As String
      Case "+"
         S2 = "IMSUM"
      Case "-"
         S2 = "IMSUB"
      Case "*"
         S2 = "IMPRODUCT"
      Case "/"
         S2 = "IMDIV"
      Case Else
         S2 = S
   End Select
   FINDSYMBOL = S2 & "("
End Function
