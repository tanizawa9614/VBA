Attribute VB_Name = "S_数式設定"
Option Explicit

Sub タブの追加()
Attribute タブの追加.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
    Dim a As Object
    Set a = ActiveDocument.PageSetup
    With ActiveDocument.PageSetup
        Selection.ParagraphFormat.TabStops.Add Position:=.PageWidth / 2 - .LeftMargin, Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces
        Selection.ParagraphFormat.TabStops.Add Position:=.PageWidth - .RightMargin - .LeftMargin, Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
    End With
    Selection.TypeText Text:=vbTab & vbTab & "()"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "SEQ eq\* Arabic", PreserveFormatting:=False
End Sub

Sub フィールドの自動更新()
   Dim aStory As Range
   Dim aField As Field

   For Each aStory In ActiveDocument.StoryRanges
      For Each aField In aStory.Fields
         aField.Update
      Next
   Next

End Sub

