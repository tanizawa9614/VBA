Attribute VB_Name = "S_10分雨量"
Option Explicit

Sub S_10分雨量()
   Dim Target_t As Range, Target_R As Range
   Dim write_log As Range
   On Error GoTo myErr
   Const msg = "選択する範囲は縦方向の単一列でなければなりません" & vbCr
   Set Target_t = Application.InputBox _
         ("洪水到達時間 t に該当するセル範囲を選択してください" _
         & vbCr & vbCr & "注)" & msg, Type:=8, Title:="【セル範囲の選択 1/3】 洪水到達時間 t")
   Set Target_R = Application.InputBox _
         ("洪水到達時間雨量 R に該当するセル範囲を選択してください" _
         & vbCr & vbCr & "注)" & msg, Type:=8, Title:="【セル範囲の選択 2/3】 洪水到達時間雨量 R")
   Set write_log = Application.InputBox _
         ("結果の出力先セルを選択してください" & vbCr, _
         Type:=8, Title:="【セル範囲の選択 3/3】 結果の出力先")
   Dim t, R
   t = Target_t.Value
   R = Target_R.Value
   If UBound(t, 1) = 1 Or UBound(R, 1) = 1 Or UBound(t, 2) >= 2 Or UBound(R, 2) >= 2 Then
'      t = WorksheetFunction.Transpose(t) 'おそらくこの関数は使えないので
'      R = WorksheetFunction.Transpose(R)
      MsgBox "tおよびRの" & msg
      End
   End If
   
   Dim n As Long, i As Long, j As Long
   Dim t2(), t2_10(), sta_n(), end_n()
   Dim row_sum As Long

   n = UBound(t, 1)
   ReDim t2(n - 1), t2_10(n - 1)
   ReDim sta_n(n - 1), end_n(n - 1)

   sta_n(0) = 0
   For i = 0 To n - 1
      t2(i) = t(i + 1, 1) - sta_n(i)
      t2_10(i) = Int(t2(i) / 10)
      end_n(i) = t2(i) Mod 10
      If i <> n - 1 Then sta_n(i + 1) = 10 - end_n(i)
      row_sum = row_sum + t2_10(i)
      If end_n(i) <> 0 Then row_sum = row_sum + 1
   Next i
   
   Dim buf(), cnt As Long, row_count As Long
   ReDim buf(row_sum - 1, 3 * n - 1)
   For j = 0 To 3 * n - 1
      If ((j + 1) + 1) Mod 3 = 0 Then
         For cnt = 1 To t2_10(Int(j / 3))
            buf(row_count, j) = 10
            row_count = row_count + 1
         Next
      ElseIf ((j + 1) + 1) Mod 3 = 2 Then
         buf(row_count, j) = sta_n(Int(j / 3))
         If sta_n(Int(j / 3)) <> 0 Then row_count = row_count + 1
      Else
         If end_n(Int(j / 3)) = 0 Then row_count = row_count - 1
         buf(row_count, j) = end_n(Int(j / 3))
      End If
   Next
   Dim ans()
   ReDim ans(row_sum - 1, 0)
   For i = 0 To UBound(buf, 1)
      For j = 0 To UBound(buf, 2)
         ans(i, 0) = ans(i, 0) + buf(i, j) * R(Int(j / 3) + 1, 1) / t(Int(j / 3) + 1, 1)
      Next
   Next
myErr:
   Const Errmsg = vbCr & "① セル範囲を選択しなかった．" & vbCr _
   & "② セル以外のものを入力した" & vbCr & "③ その他(開発者へ相談)"
   If Err.Number > 0 Then
      MsgBox "エラーが発生しました.考えられる原因は以下の通りです．" & Errmsg
      Exit Sub
   End If
   Dim flag
   If WorksheetFunction.CountA(write_log.Resize(row_sum)) >= 1 Then
      flag = MsgBox("結果を出力しようとしているセルの範囲内(下方)に既存のデータが存在しています．" _
      & vbCr & "結果を出力しますか？", vbYesNo + vbQuestion)
      If flag = vbNo Then
         MsgBox "処理を中断しました"
         Exit Sub
      End If
   End If
'   write_log.Resize(row_sum, 3 * n).Offset(, 1) = buf
   write_log.Resize(row_sum) = ans
End Sub
