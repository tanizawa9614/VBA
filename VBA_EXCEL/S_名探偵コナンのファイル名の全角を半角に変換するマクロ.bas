Attribute VB_Name = "Module1"
Option Explicit

Sub 数字を半角に()
  Dim F As String, N As String, NN As String
  Dim Place As String
  Place = "E:\3名探偵コナン\"
  F = Dir(Place & "*.*")
  Do While F <> ""
    If InStr(F, "話") <> 0 Then
        N = StrConv(Mid(F, 2, InStr(F, "話") - 1), vbWide)
        NN = Left(F, 1) & N & Mid(F, InStr(F, "話") + 1, Len(F))
        Name Place & F As Place & NN
    End If
    F = Dir()
  Loop
End Sub
