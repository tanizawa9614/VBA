Attribute VB_Name = "S_�f��������"
Option Explicit

Sub �f��������() 'Main�}�N��
   Dim i As Long, j As Long, k As Long
   Dim myA, A As String
   Dim n As Long
   Dim i_max As Long, cnt As Long
   n = 16001
   i_max = Int(Sqr(n))
   If n > 8 Then
      ReDim myA(i_max)
   Else
      ReDim myA(2)
   End If
   For i = 2 To i_max
      If n < i Then Exit For
      If n Mod i = 0 Then
         If i�͑f���ł���(i) Then
            cnt = ���񊄂�邩(n, i)
            n = n / i ^ cnt
            If A <> "" Then
               A = A & "*" & i & "^" & cnt
            Else
               A = i & "^" & cnt
            End If
            i_max = Int(Sqr(n))
'            For j = 1 To cnt
'               myA(k) = i
'               k = k + 1
'            Next j
         End If
      End If
   Next i
   If n = 1 Then
      ReDim Preserve myA(k - 1)
   Else
      myA(k) = n
      ReDim Preserve myA(k)
      If A <> "" Then
         A = A & "*" & n
      Else
         A = n
      End If
   End If
   MsgBox A
'   MsgBox Join(myA)
End Sub

Function ���񊄂�邩(ByVal �폜�� As Long, ���� As Long)
   Dim cnt As Long
   Do
      If �폜�� Mod ���� <> 0 Then
         ���񊄂�邩 = cnt
         Exit Function
      End If
      �폜�� = �폜�� / ����
      cnt = cnt + 1
   Loop
End Function

Function i�͑f���ł���(n1 As Long) As Boolean
   Dim i1 As Long, i1_max As Long
   i1_max = Int(Sqr(n1))
   For i1 = 2 To i1_max
      If n1 Mod i1 = 0 Then
         i�͑f���ł��� = False
         Exit Function
      End If
   Next
   i�͑f���ł��� = True
End Function
