Attribute VB_Name = "S_�t�B�[���h�̎����X�V"
Option Explicit

Sub �t�B�[���h�̎����X�V()
   Dim aStory As Range
   Dim aField As Field

   For Each aStory In ActiveDocument.StoryRanges
      For Each aField In aStory.Fields
         aField.Update
      Next
   Next

End Sub
