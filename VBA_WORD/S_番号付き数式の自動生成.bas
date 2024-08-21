Attribute VB_Name = "S_”Ô†•t‚«”®‚Ì©“®¶¬"
Option Explicit

Sub ”Ô†•t‚«”®‚Ì©“®¶¬()
Attribute ”Ô†•t‚«”®‚Ì©“®¶¬.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
    Application.Templates( _
        "C:\Users\yuuki\AppData\Roaming\Microsoft\Templates\Normal.dotm"). _
        BuildingBlockEntries("user_eqn").Insert Where:=Selection.Range, RichText:=True
End Sub
