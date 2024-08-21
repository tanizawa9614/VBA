Attribute VB_Name = "S_FORMULA_Buildup"
Option Explicit
Sub FORMULA_Buildup()

    Dim InitRange As Word.Range
    Dim objRange As Word.Range
    Dim objOMFun As Word.OMathFunction
    Dim objSEL As Word.Selection
    
    Set InitRange = Selection.Range
    Set objRange = Selection.OMaths.Add(InitRange)

'    Set objOMFun = objRange.OMaths(1).Functions.Add(objRange, wdOMathFunctionMat)
    Set objSEL = Selection  '文字位置取得
    
    '1文字右へ移動して数式セット
'    objSEL.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="(\sigma _x+\delta \sigma _x)"
    Set objSEL = Selection  '文字位置取得
    
    '数式の画面編集
    objRange.OMaths(1).BuildUp
    
End Sub
Sub Test_Sample_Miniature()

    Dim InitRange As Word.Range
    Dim objRange As Word.Range
    Dim objOMFun As Word.OMathFunction
    Dim objSEL As Word.Selection
    
    '末尾へ数式オブジェクトを設定
    ActiveDocument.Bookmarks("\EndOfDoc").Select
    Selection.TypeParagraph
    Set InitRange = Selection.Range
    InitRange.Text = " "
    Set objRange = Selection.OMaths.Add(InitRange)

    '∑式を追加
    Set objOMFun = objRange.OMaths(1).Functions.Add(objRange, wdOMathFunctionNary)
    Set objSEL = Selection  '文字位置取得
    
    objOMFun.Nary.Char = 8721
    objOMFun.Nary.HideSub = True
    objOMFun.Nary.HideSup = True
    
    '1文字右へ移動して数式セット
    objSEL.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="a/b"
    Set objSEL = Selection  '文字位置取得
    
    '1文字右へ移動して数式セット
    objSEL.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="=x"
    Set objSEL = Selection  '文字位置取得
    
    '数式の画面編集
    objRange.OMaths(1).BuildUp
    
    '√式を追加
    Set objOMFun = Selection.OMaths(1).Functions.Add(Selection.Range, wdOMathFunctionRad)
    objOMFun.Rad.HideDeg = True
    Set objSEL = Selection  '文字位置取得
    
    '1文字左へ移動してルート内へ数式セット
    objSEL.MoveLeft Unit:=wdCharacter, Count:=1
    objSEL.TypeText Text:="x+1"
    Set objSEL = Selection  '文字位置取得
    
    '1文字左へ移動して分母数式セット
    objSEL.MoveRight Unit:=wdCharacter, Count:=1
    objSEL.TypeText Text:="/(a+b)"
    
    '数式の画面編集
    objRange.OMaths(1).BuildUp
    
    '移動100文字で欄外へぬける。
    On Error Resume Next
    Selection.MoveRight Unit:=wdCharacter, Count:=100
    
End Sub

