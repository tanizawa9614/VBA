Attribute VB_Name = "S_�Q�Ɛݒ�"
Option Explicit

Sub �Q�Ɛݒ�()
   Dim Ref, buf, flag As Boolean
   Const RefFile As String = "C:\Windows\SysWOW64\scrrun.dll"
   Const SR As String = "Microsoft Scripting Runtime"
   For Each Ref In ActiveWorkbook.VBProject.References
      If Ref.Description = SR Then
         flag = False
         Exit For
      End If
      flag = True
   Next
   If flag Then
      ActiveWorkbook.VBProject.References.AddFromFile RefFile
   End If
End Sub
