Attribute VB_Name = "S_mesh"
Option Explicit

Sub mesh()
Attribute mesh.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim a As Double
    a = Application.CentimetersToPoints(1)
    Selection.ColumnWidth = 4.1
    Selection.RowHeight = a
    
End Sub
