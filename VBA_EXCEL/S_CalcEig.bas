Attribute VB_Name = "S_CalcEig"
Option Explicit
Option Base 1 ' 配列のインデックスを1から始める

Sub CalculateEigenvaluesAndEigenvectors()
    Dim n As Long
    Dim a() As Double, q() As Double, r() As Double, eigenvalues() As Double, eigenvectors() As Double
    Dim i As Long, j As Long
    
    n = 3 ' 行列の次元
    
    ReDim a(1 To n, 1 To n)
    ReDim q(1 To n, 1 To n)
    ReDim r(1 To n, 1 To n)
    ReDim eigenvalues(1 To n)
    ReDim eigenvectors(1 To n, 1 To n)
    
    ' 行列Aの要素を設定（ここでは適当な値を使用）
    a(1, 1) = 2: a(1, 2) = -1: a(1, 3) = 0
    a(2, 1) = -1: a(2, 2) = 2: a(2, 3) = -1
    a(3, 1) = 0: a(3, 2) = -1: a(3, 3) = 2
    
    ' 行列AのQR分解を計算
    Call QRDecomposition(a, q, r)
    
    ' QR分解から固有値と固有ベクトルを抽出
    For i = 1 To n
        eigenvalues(i) = r(i, i)
        For j = 1 To n
            eigenvectors(i, j) = q(j, i)
        Next j
    Next i
    
    ' 結果を表示
    Debug.Print "Eigenvalues:"
    For i = 1 To n
        Debug.Print eigenvalues(i)
    Next i
    
    Debug.Print "Eigenvectors:"
    For i = 1 To n
        For j = 1 To n
            Debug.Print eigenvectors(i, j)
        Next j
    Next i
End Sub

Sub QRDecomposition(a() As Double, q() As Double, r() As Double)
    Dim n As Long
    Dim i As Long, j As Long, k As Long
    Dim v() As Double, u() As Double
    Dim norm As Double, alpha As Double
    
    n = UBound(a, 1)
    ReDim v(1 To n)
    ReDim u(1 To n)
    
    ' 行列Aのコピーを作成
    For i = 1 To n
        For j = 1 To n
            q(i, j) = a(i, j)
        Next j
    Next i
    
    ' 行列Qを単位行列に初期化
    For i = 1 To n
        For j = 1 To n
            If i = j Then
                r(i, j) = 1
            Else
                r(i, j) = 0
            End If
        Next j
    Next i
    
    ' Householder変換によるQR分解
    For k = 1 To n - 1
        norm = 0
        For i = k To n
            norm = norm + q(i, k) ^ 2
        Next i
        norm = Sqr(norm)
        
        If q(k, k) >= 0 Then
            alpha = -norm
        Else
            alpha = norm
        End If
        
        For i = 1 To n
            u(i) = 0
        Next i
        
        u(k) = q(k, k) - alpha
        For i = k + 1 To n
            u(i) = q(i, k)
        Next i
        
        norm = 0
        For i = k To n
            norm = norm + u(i) ^ 2
        Next i
        norm = Sqr(norm)
        
        If norm = 0 Then
            Exit For
        End If
        
        For i = k To n
            u(i) = u(i) / norm
        Next i
        
        For j = k To n
            r(k, j) = 0
            For i = k To n
                r(k, j) = r(k, j) + 2 * u(i) * q(i, j)
            Next i
        Next j
        
        For j = 1 To n
            For i = k To n
                q(i, j) = q(i, j) - r(k, j) * u(i)
            Next i
        Next j
    Next k
End Sub

