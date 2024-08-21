Attribute VB_Name = "F_LeastSquares"
Option Explicit

Function LeastSquares(���m��Y As Range, ���m��X As Range, Optional includeIntercept As Boolean = True) As Variant
    ' ����:
    ' ���m��X: �����ϐ��̃f�[�^���i�[���ꂽ�Z���͈�
    ' ���m��Y: �]���ϐ��̃f�[�^���i�[���ꂽ�Z���͈�
    ' includeIntercept: �萔���ib�j���܂߂邩�ǂ����̃t���O�iTrue: �܂߂�AFalse: �܂߂Ȃ��j

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

    ' �f�[�^�̎擾
    dataX = ���m��X.Value
    dataY = ���m��Y.Value
    numRows = ���m��X.rows.count
    numCols = ���m��X.Columns.count

    ' H �s��̍쐬
    numCoefficients = numCols
    If includeIntercept Then numCoefficients = numCoefficients + 1 ' �萔�����܂߂�ꍇ�͗�+1�A�܂߂Ȃ��ꍇ�͗񐔂ƂȂ�
    ReDim hMatrix(1 To numRows, 1 To numCoefficients)
    For i = 1 To numRows
        For j = 1 To numCols
            hMatrix(i, j) = dataX(i, j)
        Next j
        If includeIntercept Then
            hMatrix(i, numCoefficients) = 1 ' �萔���̗�
        End If
    Next i

    ' y �x�N�g���̍쐬
    ReDim yVector(1 To numRows, 1 To 1)
    For i = 1 To numRows
        yVector(i, 1) = dataY(i, 1)
    Next i

    ' H �s��̓]�u
    transposeH = MatrixTranspose(hMatrix)

    ' H^T * H �s��̌v�Z
    hTransposeH = MatrixMultiplication(transposeH, hMatrix)

    ' H^T * y �x�N�g���̌v�Z
    hTransposeY = MatrixMultiplication(transposeH, yVector)

    ' �W���̌v�Z
    coefficients = MatrixMultiplication(MatrixInverse(hTransposeH), hTransposeY)

    ' ���ʂ�Ԃ�
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
    
    ' �P�ʍs����쐬
    For i = 1 To size
        For j = 1 To size
            If i = j Then
                matrixB(i, j) = 1
            Else
                matrixB(i, j) = 0
            End If
        Next j
    Next i
    
    ' �s��̃R�s�[
    For i = 1 To size
        For j = 1 To size
            matrixC(i, j) = matrixA(i, j)
        Next j
    Next i
    
    ' �K�E�X�E�W�����_���@�ɂ��t�s��̌v�Z
    For i = 1 To size
        temp = matrixC(i, i)
        
        ' ��Ίp�v�f��1�ɂ���
        For j = 1 To size
            matrixC(i, j) = matrixC(i, j) / temp
            matrixB(i, j) = matrixB(i, j) / temp
        Next j
        
        ' ��Ίp�v�f�ȊO��0�ɂ���
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
