Attribute VB_Name = "F_MODCAL"
Option Explicit

Function MODCAL(������ As Long, _
                  Optional �w���� As Long = 1, _
                     Optional �\������ As Long = 2)
   Dim n As Long, n0 As Long, i As Long
   n0 = ������ Mod 10 ^ �\������
   n = n0
   For i = 1 To �w����
      If i <> 1 Then n = (n0 * n) Mod 10 ^ �\������
   Next i
   MODCAL = n
End Function
