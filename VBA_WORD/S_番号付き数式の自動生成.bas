Attribute VB_Name = "S_�ԍ��t�������̎�������"
Option Explicit

Sub �ԍ��t�������̎�������()
Attribute �ԍ��t�������̎�������.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
    Application.Templates( _
        "C:\Users\yuuki\AppData\Roaming\Microsoft\Templates\Normal.dotm"). _
        BuildingBlockEntries("user_eqn").Insert Where:=Selection.Range, RichText:=True
End Sub
