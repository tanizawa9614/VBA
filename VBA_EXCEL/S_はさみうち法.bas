Attribute VB_Name = "S_�͂��݂����@"
Option Explicit

Sub �͂��݂����@()
    Dim x As Double, a As Double, b As Double, c As Double
    Dim fa As Double, fb As Double, fc As Double
    Dim eps As Double, cnt As Long
    a = 10
    b = -10
    fa = f(a)
    fb = f(b)
    c = a - fa * (a - b) / (fa - fb)
    fc = f(c)
    eps = Abs(fc)
    Do While eps > 10 ^ -10
        If fc < 0 Then
            If fa < 0 Then
                a = c
            Else
                b = c
            End If
        Else
           If fa > 0 Then
                a = c
            Else
                b = c
            End If
        End If
        fa = f(a)
        fb = f(b)
        c = a - fa * (a - b) / (fa - fb)
        fc = f(c)
        eps = Abs(fc)
        If cnt > 10 ^ 5 Then Stop
        cnt = cnt + 1
    Loop
    MsgBox "�ߎ��� :  x =  " & c & vbCr & "���̎��̊֐��l :  f(x) = " & fc & vbCr & "���[�v�� : " & cnt - 1
End Sub
Function f(x As Double)
    f = (x - 1) * (x - 2) * (x - 3) * (x - 4)
End Function
