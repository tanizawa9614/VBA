Attribute VB_Name = "S_全ファイル名取得_セルから"
Option Explicit

Sub 全ファイル名取得_Pathはセルから()
    Dim FolPath As String, tmpArr, FolArr
    Dim A()
    Dim pfol As Object
    Dim n As Long
    Dim FSO As Object
    Dim sh As Worksheet
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set sh = Worksheets("FolderPath")
    tmpArr = sh.Range("A1").CurrentRegion
    FolArr = Convert2Dim(tmpArr)
    
    Dim i As Long, j As Long, cnt As Long
    n = 0
    cnt = 0
    For i = 1 To UBound(FolArr)
        FolPath = FolArr(i, 1)
        Set pfol = FSO.GetFolder(FolPath)
        
        Call FilesCount(pfol, n)
        If i = 1 Then
            ReDim A(1 To n)
        Else
            ReDim Preserve A(1 To n)
        End If
                
        Call GetFiles(A, pfol, cnt)
                
    Next
        
    sh.Columns(3).ClearContents
    sh.Cells(1, 3).Resize(n, 1) = WorksheetFunction.Transpose(A)
End Sub

Private Sub FilesCount(ByVal pfol As Object, ByRef n As Long)
    Dim subfol As Object
    For Each subfol In pfol.subfolders
        Call FilesCount(subfol, n)
    Next
    n = n + pfol.Files.Count + 1
End Sub

Private Sub GetFiles(ByRef A, ByVal pfol As Object, ByRef n As Long)
    Dim subfol As Object
    Dim subfile As Object
    Dim i As Long
    
    For Each subfol In pfol.subfolders
        Call GetFiles(A, subfol, n)
    Next
    n = n + 1
    A(n) = pfol.path
    For Each subfile In pfol.Files
        n = n + 1
        A(n) = subfile.path
    Next
End Sub

Private Function Convert2Dim(arr0)
    Dim ans() As Variant, arr
    Dim i As Long, j As Long
    Dim ic As Long, jc As Long
    Dim numRows As Long, numCols As Long
    
    arr = arr0
    
    If IsArray(arr) Then
        On Error Resume Next
        numRows = UBound(arr, 1) - LBound(arr, 1) + 1
        numCols = UBound(arr, 2) - LBound(arr, 2) + 1
        On Error GoTo 0
        
        If numCols = 0 Then ' 1次元配列の場合
            ReDim ans(1 To 1, 1 To numRows)
            
            For i = LBound(arr) To UBound(arr)
                ic = ic + 1
                ans(1, ic) = arr(i)
            Next i
            
        Else ' 2次元配列の場合
            If LBound(arr, 1) * LBound(arr, 2) = 1 Then
                ans = arr
            Else
                ReDim ans(1 To numRows, 1 To numCols)
                For i = LBound(arr, 1) To UBound(arr, 1)
                    ic = ic + 1
                    jc = 0
                    For j = LBound(arr, 2) To UBound(arr, 2)
                        jc = jc + 1
                        ans(ic, jc) = arr(i, j)
                    Next j
                Next i
            End If
        End If
    Else
        ' If input is not an array, create a 1x1 array with minimum index 1
        ReDim ans(1 To 1, 1 To 1)
        ans(1, 1) = arr
    End If
    
    Convert2Dim = ans
End Function

