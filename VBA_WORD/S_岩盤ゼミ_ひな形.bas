Attribute VB_Name = "S_��Ճ[�~_�ЂȌ`"
Option Explicit

Sub ��Ճ[�~_�ЂȌ`()
Attribute ��Ճ[�~_�ЂȌ`.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro3"
   With Selection.Font
      .NameFarEast = "�l�r ����"
      .NameAscii = "Times New Roman"
      .NameOther = "Times New Roman"
      .Size = 10
   End With
   ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
   With Selection
      .Font.Name = "�l�r �S�V�b�N"
      .Font.Name = "Times New Roman"
      .Font.Size = 13
      .ParagraphFormat.LineSpacing = LinesToPoints(1.5)
      .ParagraphFormat.Alignment = wdAlignParagraphCenter
      .TypeText Text:="��Ճ[�~"
      .TypeParagraph
      .ParagraphFormat.Alignment = wdAlignParagraphRight
      .TypeText Text:=Replace(Format(Now, "yyyy/mm/dd"), "/", ".")
      .TypeParagraph
      .TypeText Text:="�J�V �Ў�"
   End With
   
   ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
   With Selection
      .Font.Name = "�l�r �S�V�b�N"
      .Font.Name = "Times New Roman"
      .Font.Size = 10.5
      .ParagraphFormat.LineSpacing = LinesToPoints(1.5)
   End With
   Application.Templates( _
      "C:\Users\yuuki\AppData\Roaming\Microsoft\Document Building Blocks\1041\16\Built-In Building Blocks.dotx" _
      ).BuildingBlockEntries("�ԍ��̂� 2").Insert Where:=Selection.Range, RichText _
      :=True
   ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

End Sub
