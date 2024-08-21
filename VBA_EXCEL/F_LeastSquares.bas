Attribute VB_Name = "F_LeastSquares"
Option Explicit

Function LeastSquares(既知のY As Range, 既知のX As Range, Optional includeIntercept As Boolean = True) As Variant
    ' 引数:
    ' 既知のX: 自立変数のデータが格納されたセル範囲
    ' 既知のY: 従属変数のデータが格納されたセル範囲
    ' includeIntercept: 定数項（b）を含めるかどうかのフラグ（True: 含める、False: 含めない）

    Dim dataX As Variant
    Dim dataY As Variant
    Dim numRows As Long
    Dim numCols As Long
    Dim numCoefficients As Long
    Dim coefficients As Variant
    Dim hMatrix As Variant
    Dim yVector As Variant
    Dim transposeH As Variant
    Dim hTransposeH As Variant
    Dim hTransposeY As Variant
    Dim i As Long
    Dim j As Long

    ' データの取得
    dataX = 既知のX.Value
    dataY = 既知のY.Value
    numRows = 既知のX.rows.count
    numCols = 既知のX.Columns.count

    ' H 行列の作成
    numCoefficients = numCols
    If includeIntercept Then numCoefficients = numCoefficients + 1 ' 定数項を含める場合は列数+1、含めない場合は列数となる
    ReDim hMatrix(1 To numRows, 1 To numCoefficients)
    For i = 1 To numRows
        For j = 1 To numCols
            hMatrix(i, j) = dataX(i, j)
        Next j
        If includeIntercept Then
            hMatrix(i, numCoefficients) = 1 ' 定数項の列
        End If
    Next i

    ' y ベクトルの作成
    ReDim yVector(1 To numRows, 1 To 1)
    For i = 1 To numRows
        yVector(i, 1) = dataY(i, 1)
    Next i

    ' H 行列の転置
    transposeH = MatrixTranspose(hMatrix)

    ' H^T * H 行列の計算
    hTransposeH = MatrixMultiplication(transposeH, hMatrix)

    ' H^T * y ベクトルの計算
    hTransposeY = MatrixMultiplication(transposeH, yVector)

    ' 係数の計算
    coefficients = MatrixMultiplication(MatrixInverse(hTransposeH), hTransposeY)

    ' 結果を返す
    LeastSquares = coefficients
End Function

Private Function MatrixInverse(matrixA As Variant) As Variant
    Dim size As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim temp As Double
    Dim determinant As Double
    Dim matrixB() As Double
    Dim matrixC() As Double
    
    size = UBound(matrixA, 1)
    ReDim matrixB(1 To size, 1 To size)
    ReDim matrixC(1 To size, 1 To size)
    
    ' 単位行列を作成
    For i = 1 To size
        For j = 1 To size
            If i = j Then
                matrixB(i, j) = 1
            Else
                matrixB(i, j) = 0
            End If
        Next j
    Next i
    
    ' 行列のコピー
    For i = 1 To size
        For j = 1 To size
            matrixC(i, j) = matrixA(i, j)
        Next j
    Next i
    
    ' ガウス・ジョルダン法による逆行列の計算
    For i = 1 To size
        temp = matrixC(i, i)
        
        ' 主対角要素を1にする
        For j = 1 To size
            matrixC(i, j) = matrixC(i, j) / temp
            matrixB(i, j) = matrixB(i, j) / temp
        Next j
        
        ' 主対角要素以外を0にする
        For j = 1 To size
            If j <> i Then
                temp = matrixC(j, i)
                
                For k = 1 To size
                    matrixC(j, k) = matrixC(j, k) - temp * matrixC(i, k)
                    matrixB(j, k) = matrixB(j, k) - temp * matrixB(i, k)
                Next k
            End If
        Next j
    Next i
    
    MatrixInverse = matrixB
End Function
Private Function MatrixMultiplication(matrixA As Variant, matrixB As Variant) As Variant
    Dim rowsA As Integer
    Dim colsA As Integer
    Dim colsB As Integer
    Dim resultMatrix() As Double
    Dim i As Integer, j As Integer, k As Integer
    Dim temp As Double
    
    rowsA = UBound(matrixA, 1)
    colsA = UBound(matrixA, 2)
    colsB = UBound(matrixB, 2)
    
    ReDim resultMatrix(1 To rowsA, 1 To colsB)
    
    For i = 1 To rowsA
        For j = 1 To colsB
            temp = 0
            For k = 1 To colsA
                temp = temp + matrixA(i, k) * matrixB(k, j)
            Next k
            resultMatrix(i, j) = temp
        Next j
    Next i
    
    MatrixMultiplication = resultMatrix
End Function

Private Function MatrixTranspose(matrix As Variant) As Variant
    Dim rows As Integer
    Dim cols As Integer
    Dim transposedMatrix() As Double
    Dim i As Integer, j As Integer
    
    rows = UBound(matrix, 2)
    cols = UBound(matrix, 1)
    
    ReDim transposedMatrix(1 To rows, 1 To cols)
    
    For i = 1 To rows
        For j = 1 To cols
            transposedMatrix(i, j) = matrix(j, i)
        Next j
    Next i
    
    MatrixTranspose = transposedMatrix
End Function
