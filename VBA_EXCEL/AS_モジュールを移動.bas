Attribute VB_Name = "AS_���W���[�����ړ�"
Option Explicit
Dim ex_wb As Workbook
Dim im_wb As Workbook
Dim FSO As Object
Dim ex_modu As Object

Sub ���W���[�����ړ�()
   Set ex_wb = Workbooks("General.xlam")
   Set im_wb = ActiveWorkbook
   Set FSO = CreateObject("Scripting.FileSystemObject")
   
   For Each ex_modu In ex_wb.VBProject.VBComponents
      If ex_modu.Name <> "Sheet1" Then
         If ex_modu.Name = "ThisWorkbook" Then
            Call Workbook���W���[���̏ꍇ
         Else
            Call �W�����W���[���̏ꍇ
         End If
      End If
   Next
   Call Book���A�h�C���Ƃ��ĕۑ�
End Sub

Private Sub �W�����W���[���̏ꍇ()
   Const temp_place = "C:\Users\yuuki\AppData\Roaming\Microsoft\AddIns"
   Dim temp_file As String, ext As String
   ext = ".bas"
   temp_file = temp_place & ex_modu.Name & ext
   ex_modu.Export temp_file
   im_wb.VBProject.VBComponents.Import temp_file
   FSO.GetFile(temp_file).Delete
End Sub

Private Sub Workbook���W���[���̏ꍇ()
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

Private Sub Book���A�h�C���Ƃ��ĕۑ�()
   im_wb.SaveAs _
      "C:\Users\yuuki\AppData\Roaming\Microsoft\AddIns\General2.xlam", _
      FileFormat:=xlOpenXMLAddIn
   Shell "C:\Windows\Explorer.exe " & _
      "C:\Users\yuuki\AppData\Roaming\Microsoft\AddIns", vbNormalFocus
End Sub
