Attribute VB_Name = "S_拡張メタファイルを貼付け"
Option Explicit

Sub 拡張メタファイルを貼付け()
    On Error Resume Next
    ActiveWindow.Selection.ShapeRange(1).Copy  '選択している図形がある場合はそれをコピー，ない場合はコピーされている図形
    ActiveWindow.View.PasteSpecial DataType:=ppPasteEnhancedMetafile  '拡張メタファイルとして貼付け
    On Error GoTo 0
End Sub

