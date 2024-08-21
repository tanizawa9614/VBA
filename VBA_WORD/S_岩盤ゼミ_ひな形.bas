Attribute VB_Name = "S_ä‚î’É[É~_Ç–Ç»å`"
Option Explicit

Sub ä‚î’É[É~_Ç–Ç»å`()
Attribute ä‚î’É[É~_Ç–Ç»å`.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro3"
   With Selection.Font
      .NameFarEast = "ÇlÇr ñæí©"
      .NameAscii = "Times New Roman"
      .NameOther = "Times New Roman"
      .Size = 10
   End With
   ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
   With Selection
      .Font.Name = "ÇlÇr ÉSÉVÉbÉN"
      .Font.Name = "Times New Roman"
      .Font.Size = 13
      .ParagraphFormat.LineSpacing = LinesToPoints(1.5)
      .ParagraphFormat.Alignment = wdAlignParagraphCenter
      .TypeText Text:="ä‚î’É[É~"
      .TypeParagraph
      .ParagraphFormat.Alignment = wdAlignParagraphRight
      .TypeText Text:=Replace(Format(Now, "yyyy/mm/dd"), "/", ".")
      .TypeParagraph
      .TypeText Text:="íJ‡V ò–é˜"
   End With
   
   ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
   With Selection
      .Font.Name = "ÇlÇr ÉSÉVÉbÉN"
      .Font.Name = "Times New Roman"
      .Font.Size = 10.5
      .ParagraphFormat.LineSpacing = LinesToPoints(1.5)
   End With
   Application.Templates( _
      "C:\Users\yuuki\AppData\Roaming\Microsoft\Document Building Blocks\1041\16\Built-In Building Blocks.dotx" _
      ).BuildingBlockEntries("î‘çÜÇÃÇ› 2").Insert Where:=Selection.Range, RichText _
      :=True
   ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

End Sub
