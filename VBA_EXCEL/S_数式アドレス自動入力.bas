Attribute VB_Name = "S_数式アドレス自動入力"
Option Explicit
Dim buf

Sub 数式アドレス自動入力()
   Dim A, i As Long, j As Long
   Dim fmla As Range, buf2, buf3
   Dim xTarget As Range, Target As Range
   
   On Error GoTo myErr
   Set fmla = Application.InputBox("数式を入力したセルを選択してください", Type:=8)
   Set Target = ActiveCell
'   Set Target = Application.InputBox("入力先のセルを選択してください", Type:=8)
   buf = 数式のフォーマット調整(fmla.Value)
   
   A = 引数把握
   For i = LBound(A) To UBound(A)
      buf3 = ""
      Set xTarget = Application.InputBox(A(i) & "を入力してください", Type:=8)
      buf2 = Split(buf, A(i))
      For j = LBound(buf2) To UBound(buf2)
         If 引数である(buf2, j) = True Then
            Select Case xTarget.HasSpill
            Case True
               buf3 = buf3 & buf2(j) & xTarget.Resize(1, 1).Address(False, False) & "#"
            Case False
               buf3 = buf3 & buf2(j) & xTarget.Address(False, False)
            End Select
         ElseIf j < UBound(buf2) - 1 Then
            buf3 = buf3 & buf2(j) & A(i)
         Else
            buf3 = buf3 & buf2(j)
         End If
      Next
      buf = buf3
   Next
   Target.Formula2 = "=" & buf
myErr:
End Sub
Private Function 引数である(A, cnt As Long) As Boolean
   Dim Cst2(5), i As Long
   Cst2(0) = "+": Cst2(1) = "-": Cst2(2) = "*"
   Cst2(3) = "/": Cst2(4) = "^": Cst2(5) = ")"
   If cnt = UBound(A) Then
      If A(cnt) = "" Then
         引数である = True
      Else
         引数である = False
      End If
      Exit Function
   End If
   For i = LBound(Cst2) To UBound(Cst2)
      If Cst2(i) = Left(A(cnt + 1), 1) Then
         引数である = True
         Exit Function
      End If
   Next
   引数である = False
End Function

Private Function 数式のフォーマット調整(Fun As String)
   Fun = StrConv(Fun, vbLowerCase)
   Fun = StrConv(Fun, vbNarrow)
   Fun = Trim(Fun)
   数式のフォーマット調整 = Fun
End Function

Private Function 引数把握()
   Dim CSt(30), i As Long, buf2
   Dim j As Long
   CSt(0) = "+"
   CSt(1) = "-"
   CSt(2) = "*"
   CSt(3) = "/"
   CSt(4) = "^"
   CSt(5) = "("
   CSt(6) = ")"
   CSt(7) = "exp"
   CSt(8) = "log10"
   CSt(9) = "log"
   CSt(10) = "ln"
   CSt(11) = "sqrt"
   CSt(12) = "pi"
   CSt(13) = "asin"
   CSt(14) = "acos"
   CSt(15) = "atan"
   CSt(16) = "sinh"
   CSt(17) = "cosh"
   CSt(18) = "tanh"
   CSt(19) = "sin"
   CSt(20) = "cos"
   CSt(21) = "tan"
   CSt(22) = "asec"
   CSt(23) = "acsc"
   CSt(24) = "acot"
   CSt(25) = "sech"
   CSt(26) = "csch"
   CSt(27) = "coth"
   CSt(28) = "sec"
   CSt(29) = "csc"
   CSt(30) = "cot"

   buf2 = buf
   For i = LBound(CSt) To UBound(CSt)
      buf2 = Replace(buf2, CSt(i), vbTab)
   Next
   buf2 = WorksheetFunction.Unique(Split(buf2, vbTab), True)
   For i = LBound(buf2) To UBound(buf2)
      If buf2(i) <> "" And IsNumeric(buf2(i)) = False Then
         buf2(j + 1) = buf2(i)
         j = j + 1
      End If
   Next
   ReDim Preserve buf2(j - 1)
   引数把握 = buf2
End Function

