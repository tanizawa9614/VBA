Attribute VB_Name = "S_フィールドの自動更新"
Option Explicit

Sub フィールドの自動更新()
   Dim aStory As Range
   Dim aField As Field

   For Each aStory In ActiveDocument.StoryRanges
      For Each aField In aStory.Fields
         aField.Update
      Next
   Next

End Sub
