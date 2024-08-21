Attribute VB_Name = "S_ALLCHANGE"
Option Explicit
Dim cnt_D As Long, cnt_S As Long

Sub ALLCHANGE_Sub()
   Dim path As String, ext As String
   Dim FSO As Object, fl As Object, NewName As String
   Set FSO = CreateObject("Scripting.FileSystemObject")
   With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = True Then path = .SelectedItems(1)
   End With
   For Each fl In FSO.GetFolder(path).Files
      ext = "." & FSO.GetExtensionName(fl.path)
      NewName = RECOG_NAME(fl.Name)
      If NewName <> "" And _
      FSO.FileExists(path & "\" & NewName & ext) = False _
      Then fl.Name = NewName & ext
   Next fl
   Call フォルダを立ち上げるか(path)
   Set FSO = Nothing
End Sub
Sub フォルダを立ち上げるか(C3 As String)
   Const C = "C:\Windows\explorer.exe "
   Const C1 = "E:\3名探偵コナン\", C2 = "E:\5アニメ\笑ゥせぇるすまん\"
   If cnt_D <> 0 And Dir(C1) <> "" Then Shell C & C1, 1
   If cnt_S <> 0 And Dir(C2) <> "" Then Shell C & C2, 1
   Shell C & C3, 1
End Sub

Function RECOG_NAME(F_Name As String) As String
   Const C1 = "名探偵コナン", C2 = "笑ゥせぇるすまん"
   If InStr(F_Name, C1) > 0 Then
      RECOG_NAME = DETECTIVE(F_Name)
      cnt_D = cnt_D + 1
   ElseIf InStr(F_Name, C2) > 0 Then
      RECOG_NAME = SELLS(F_Name)
      cnt_S = cnt_S + 1
   Else
      RECOG_NAME = ""
   End If
End Function

Function DETECTIVE(D_Name As String) As String
   Dim T1 As String, T2 As String
   Dim S1 As Long, S2 As Long
   S1 = InStr(D_Name, "「")
   S2 = InStr(D_Name, "」")
   T1 = Mid(D_Name, S1, S2 - S1 + 1)
   S1 = InStr(D_Name, "第")
   S2 = InStr(D_Name, "話")
   T2 = Mid(D_Name, S1, S2 - S1 + 1)
   DETECTIVE = T2 & " " & T1
End Function
Function SELLS(S_Name As String) As String
   Dim T1 As String, T2 As String
   Dim S1 As Long, S2 As Long
   On Error Resume Next
   S1 = InStr(S_Name, "「")
   S2 = InStr(S_Name, "」")
   T1 = Mid(S_Name, S1 + 1, S2 - S1 - 1)
   S1 = InStr(S_Name, "（")
   S2 = InStr(S_Name, "）")
   T2 = "第" & StrConv(Mid(S_Name, S1 + 1, S2 - S1 - 1), vbNarrow) & "話 "
   If T2 = "" Then T2 = "特別編 "
   SELLS = T2 & T1
End Function

