Attribute VB_Name = "F_MODCAL"
Option Explicit

Function MODCAL(実数部 As Long, _
                  Optional 指数部 As Long = 1, _
                     Optional 表示桁数 As Long = 2)
   Dim n As Long, n0 As Long, i As Long
   n0 = 実数部 Mod 10 ^ 表示桁数
   n = n0
   For i = 1 To 指数部
      If i <> 1 Then n = (n0 * n) Mod 10 ^ 表示桁数
   Next i
   MODCAL = n
End Function
