Attribute VB_Name = "S_空フォルダの削除"
Option Explicit
Sub 空フォルダの削除()
    Dim FolPath As String, i As Long
    Dim FSO As Object, fl As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then FolPath = .SelectedItems(1)
        If FolPath = "" Then Exit Sub
    End With
    For Each fl In FSO.GetFolder(FolPath).SubFolders
        If fl.Files.Count = 0 And fl.SubFolders.Count = 0 Then
'            MsgBox fl.Name
            fl.Delete
            DoEvents
        End If
    Next
    MsgBox "終了しました"
End Sub

Sub フォルダの展開()
    Dim A As String, B As String
    Dim place As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then place = .SelectedItems(1)
        If place = "" Then Exit Sub
    End With
    
    Dim pfl As Object
    Set pfl = FSO.GetFolder(place)
    
    Dim subfl As Object
    Dim fl As Object, tmpname As String
    For Each subfl In pfl.SubFolders
        If subfl.Files.Count >= 1 Then
            For Each fl In subfl.Files
                tmpname = fl.Name
                Do While FSO.FileExists(pfl.path & "\" & tmpname)
                    tmpname = "1" & tmpname
                Loop
                If fl.Name <> tmpname Then
                    fl.Name = tmpname
                End If
                FSO.GetFile(fl.path).Move pfl.path & "\"
                DoEvents
            Next
        End If
    Next
    MsgBox "終了しました"
End Sub

