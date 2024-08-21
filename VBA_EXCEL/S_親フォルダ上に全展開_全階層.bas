Attribute VB_Name = "S_親フォルダ上に全展開_全階層"
Option Explicit
Private Sub 親フォルダ上に全展開_全階層()
    Dim FolPath As String, A()
    Dim pfol As Object
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then FolPath = .SelectedItems(1)
    End With
    
'    FolPath = "C:\Users\yuuki\OneDrive - Osaka University\デスクトップ\test"
    Set pfol = FSO.GetFolder(FolPath)
    DoEvents
    Call allfiles(A(), pfol)
    
    Dim tmpfile As Object, i As Long
    Dim tmpname As String, ext As String
    For i = 0 To UBound(A)
        If FSO.FileExists(A(i)) Then
            Set tmpfile = FSO.GetFile(A(i))
            Do While FSO.FileExists(FolPath & "\" & tmpfile.Name)
                ext = "." & FSO.GetExtensionName(A(i))
                tmpname = tmpfile.Name
                tmpname = Left(tmpname, InStrRev(tmpname, ".") - 1) & "(1)"
                tmpfile.Name = tmpname & ext
            Loop
            tmpfile.Move FolPath & "\"
        End If
    Next
    MsgBox "終了しました"
End Sub
Private Function allfiles(ByRef A(), ByVal pfol As Object) As Boolean
    Dim cnt As Long
    Dim subfol As Object
    For Each subfol In pfol.SubFolders
'        If subfol.Files.Count >= 1 Then
            Call filename(A(), subfol)
'        End If
        If subfol.SubFolders.Count >= 1 Then
            Call allfiles(A(), subfol)
        End If
    Next
End Function

Private Function filename(ByRef A(), ByVal subfol As Object)
    Dim i As Long, fl As Object
    On Error Resume Next
    i = UBound(A)
    If Err.Number > 0 Then i = -1
    On Error GoTo 0
    ReDim Preserve A(i + subfol.Files.Count + 1)
    
    i = i + 1
    A(i) = "---以下のフォルダ名；" & subfol.Name & "---ファイル数；" & subfol.Files.Count
'    A(i) = "以下のフォルダ名；" & subfol.Name
    
    For Each fl In subfol.Files
        i = i + 1
'        A(i) = fl.Path
        A(i) = fl.path
    Next
End Function
