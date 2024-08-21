Attribute VB_Name = "AS_モジュールを移動"
Option Explicit
Dim ex_wb As Workbook
Dim im_wb As Workbook
Dim FSO As Object
Dim ex_modu As Object

Sub モジュールを移動()
   Set ex_wb = Workbooks("General.xlam")
   Set im_wb = ActiveWorkbook
   Set FSO = CreateObject("Scripting.FileSystemObject")
   
   For Each ex_modu In ex_wb.VBProject.VBComponents
      If ex_modu.Name <> "Sheet1" Then
         If ex_modu.Name = "ThisWorkbook" Then
            Call Workbookモジュールの場合
         Else
            Call 標準モジュールの場合
         End If
      End If
   Next
   Call Bookをアドインとして保存
End Sub

Private Sub 標準モジュールの場合()
   Const temp_place = "C:\Users\yuuki\AppData\Roaming\Microsoft\AddIns"
   Dim temp_file As String, ext As String
   ext = ".bas"
   temp_file = temp_place & ex_modu.Name & ext
   ex_modu.Export temp_file
   im_wb.VBProject.VBComponents.Import temp_file
   FSO.GetFile(temp_file).Delete
End Sub

Private Sub Workbookモジュールの場合()
   Dim s As Long, l As Long
   Dim im_modu As Object
   
   l = ex_modu.CodeModule.CountOflines
   For Each im_modu In im_wb.VBProject.VBComponents
      If im_modu.Name = "ThisWorkbook" Then
         With im_modu.CodeModule
            .AddFromString ex_modu.CodeModule.Lines(2, l)
            Do While s + 1 < .CountOflines - 1
               If .Lines(s + 1, 1) = "" Then
                  .DeleteLines s + 1
                  s = s - 1
               End If
               s = s + 1
            Loop
         End With
         Exit For
      End If
   Next
End Sub

Private Sub Bookをアドインとして保存()
   im_wb.SaveAs _
      "C:\Users\yuuki\AppData\Roaming\Microsoft\AddIns\General2.xlam", _
      FileFormat:=xlOpenXMLAddIn
   Shell "C:\Windows\Explorer.exe " & _
      "C:\Users\yuuki\AppData\Roaming\Microsoft\AddIns", vbNormalFocus
End Sub
