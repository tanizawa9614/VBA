Attribute VB_Name = "S_番号付き数式の自動生成"
Option Explicit

Sub 番号付き数式の自動生成()
Attribute 番号付き数式の自動生成.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
    Application.Templates( _
        "C:\Users\yuuki\AppData\Roaming\Microsoft\Templates\Normal.dotm"). _
        BuildingBlockEntries("user_eqn").Insert Where:=Selection.Range, RichText:=True
End Sub
