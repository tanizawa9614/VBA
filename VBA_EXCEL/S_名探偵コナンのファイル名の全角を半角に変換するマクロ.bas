Attribute VB_Name = "Module1"
Option Explicit

Sub �����𔼊p��()
  Dim F As String, N As String, NN As String
  Dim Place As String
  Place = "E:\3���T��R�i��\"
  F = Dir(Place & "*.*")
  Do While F <> ""
    If InStr(F, "�b") <> 0 Then
        N = StrConv(Mid(F, 2, InStr(F, "�b") - 1), vbWide)
        NN = Left(F, 1) & N & Mid(F, InStr(F, "�b") + 1, Len(F))
        Name Place & F As Place & NN
    End If
    F = Dir()
  Loop
End Sub
