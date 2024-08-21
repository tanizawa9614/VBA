Attribute VB_Name = "S_ExtoPPTImageMP4"
Option Explicit

Sub ExcelからPPTに画像ファイルの貼り付けおよび動画化()
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    Dim pptShape As Object
    Dim folderPath As String
    Dim imagePath As String
    Dim imageFile As String
    Dim slideIndex As Integer
    Dim answer As Integer
    
    ' PowerPointアプリケーションを開く
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    
    ' 新しいプレゼンテーションを作成
    Set pptPres = pptApp.Presentations.Add
    
    ' スライドのサイズを設定（4:3）
    pptPres.PageSetup.SlideWidth = 914.4 ' 10インチ
    pptPres.PageSetup.SlideHeight = 685.8 ' 7.5インチ
    
    ' 画像フォルダの選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "画像フォルダを選択してください"
        .Show
        If .SelectedItems.Count > 0 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "フォルダが選択されていません。処理を終了します。", vbExclamation
            Exit Sub
        End If
    End With
    
    ' 画像フォルダ内のファイルを取得
    imageFile = Dir(folderPath & "\*.*") ' すべての拡張子に対応
    
    Do While imageFile <> ""
               
        If LCase(Right(imageFile, 4)) Like ".jpg" Or LCase(Right(imageFile, 4)) Like ".jpeg" Or _
           LCase(Right(imageFile, 4)) Like ".png" Or LCase(Right(imageFile, 4)) Like ".gif" Then
            ' スライドを追加
            slideIndex = pptPres.Slides.Count + 1
            Set pptSlide = pptPres.Slides.Add(slideIndex, 12) ' ppLayoutBlank の値: 12
            
            ' 画像をスライドに貼り付け
            imagePath = folderPath & "\" & imageFile
            Set pptShape = pptSlide.Shapes.AddPicture(Filename:=imagePath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=0, Top:=0)
            pptShape.LockAspectRatio = msoTrue
            pptShape.Top = (pptPres.PageSetup.SlideHeight - pptShape.Height) / 2
            pptShape.Left = (pptPres.PageSetup.SlideWidth - pptShape.Width) / 2
            pptShape.ScaleWidth 1, msoFalse ' 画像の幅をスライドの幅に合わせる
            pptShape.ScaleHeight 1, msoFalse ' 画像の高さをスライドの高さに合わせる
        End If
        
        ' 次の画像ファイルを取得
        imageFile = Dir
    Loop
        
    ' MP4への変換を確認
    answer = MsgBox("MP4に変換しますか？", vbYesNo + vbQuestion)
    If answer = vbYes Then
        ' プレゼンテーションをMP4として保存
        pptPres.SaveAs folderPath & "\output.mp4", 39 ' ppSaveAsMP4 の値: 39
        pptPres.Close
        MsgBox "MP4ファイルとして出力しました。処理を終了します。", vbInformation
    Else
        ' PowerPointを表示して終了
        pptApp.Visible = True
        MsgBox "処理を終了します。", vbInformation
    End If
    
    ' オブジェクトを解放
    Set pptShape = Nothing
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
End Sub

