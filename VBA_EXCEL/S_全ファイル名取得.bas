Attribute VB_Name = "S_�S�t�@�C�����擾"
Option Explicit

Sub �S�t�@�C�����擾()
    Dim FolPath As String, A()
    Dim pfol As Object
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then FolPath = .SelectedItems(1)
    End With
    
'    FolPath = "C:\Users\yuuki\OneDrive - Osaka University\�f�X�N�g�b�v\test"
    Set pfol = FSO.getfolder(FolPath)
    Call allfiles(A(), pfol)
    Call filename(A(), pfol)
    MsgBox "�I�����܂���"
    Stop
End Sub

Function allfiles(ByRef A(), ByVal pfol As Object) As Boolean
    Dim cnt As Long
    Dim subfol As Object
    For Each subfol In pfol.subfolders
'        If subfol.Files.Count >= 1 Then
            Call filename(A(), subfol)
'        End If
        If subfol.subfolders.Count >= 1 Then
            Call allfiles(A(), subfol)
        End If
    Next
End Function

Function filename(ByRef A(), ByVal subfol As Object)
    Dim i As Long, fl As Object
    On Error Resume Next
    i = UBound(A)
    If Err.Number > 0 Then i = -1
    On Error GoTo 0
    ReDim Preserve A(i + subfol.Files.Count + 1)
    
    i = i + 1
    A(i) = "---�ȉ��̃t�H���_���G" & subfol.Name & "---�t�@�C�����G" & subfol.Files.Count
'    A(i) = "�ȉ��̃t�H���_���G" & subfol.Name
    
    For Each fl In subfol.Files
        i = i + 1
'        A(i) = fl.Path
        A(i) = fl.Name
    Next
End Function
