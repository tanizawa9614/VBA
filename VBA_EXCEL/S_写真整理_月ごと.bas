Attribute VB_Name = "S_写真整理_月ごと"
Option Explicit
Sub Main1()
   Call 写真整理_月ごと
End Sub
Sub Main2()
   Call 親フォルダ上に全展開_親フォルダ指定_複数フォルダ
End Sub
Sub Main3()
   Call 空フォルダの削除
End Sub
Sub 写真整理_月ごと()
    Dim place As String, FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim fl As Object
    Dim New_Folder As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then place = .SelectedItems(1)
        If place = "" Then Exit Sub
    End With
    
'    On Error Resume Next
    Dim tmpfname As String
    For Each fl In FSO.GetFolder(place).Files
        DoEvents
        'ファイルの「作成日」を取得,新規フォルダ名は「日付」
        New_Folder = Format(FileDateTime(fl.path), "yyyym")
        New_Folder = Left(New_Folder, 4) & "年" & Mid(New_Folder, 5) & "月"
        
        If Not FSO.FolderExists(place & "\" & New_Folder) Then
            FSO.CreateFolder (place & "\" & New_Folder)
            '新規フォルダを作成
        End If
        'フォルダの「作成日」のフォルダに移動
        tmpfname = fl.Name
        Do While FSO.FileExists(place & "\" & New_Folder & "\" & tmpfname)
            tmpfname = "1" & tmpfname
        Loop
        If fl.Name <> tmpfname Then fl.Name = tmpfname
        FSO.MoveFile fl.path, place & "\" & New_Folder & "\"
    Next
    MsgBox "終了しました"
    Set FSO = Nothing
    Set fl = Nothing
End Sub

Sub 親フォルダ上に全展開_親フォルダ指定_複数フォルダ()
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
        DoEvents
        If subfl.Files.Count >= 1 Then
            For Each fl In subfl.Files
                tmpname = fl.Name
                Do While FSO.FileExists(pfl.path & "\" & tmpname)
                    tmpname = "1" & tmpname
                Loop
                If fl.Name <> tmpname Then fl.Name = tmpname
                FSO.GetFile(fl.path).Move pfl.path & "\"
            Next
        End If
    Next
    
    Call delete_emptyfile(place)
    
    MsgBox "終了しました"
    
    Set FSO = Nothing
    Set fl = Nothing
    
End Sub
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
Private Sub delete_emptyfile(Optional path As String)
    Dim flag As String, sfl As Object
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If path = "" Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = True Then path = .SelectedItems(1)
            If path = "" Then Exit Sub
        End With
    End If
    flag = MsgBox("空フォルダを削除しますか？", vbYesNo)
    If flag = vbYes Then
    For Each sfl In FSO.GetFolder(path).SubFolders
        If FSO.GetFolder(sfl.path).SubFolders.Count >= 1 Then
            MsgBox "場所：" & sfl.path & vbCr _
            & "フォルダ名：" & sfl.Name & vbCr & "にはフォルダが存在します"
            GoTo L1
        End If
        If FSO.GetFolder(sfl.path).Files.Count >= 1 Then
            MsgBox "場所：" & sfl.path & vbCr _
            & "フォルダ名：" & sfl.Name & vbCr & "にはファイルが存在します"
            GoTo L1
        End If
        FSO.DeleteFolder sfl.path
L1:
    Next
    End If
End Sub
