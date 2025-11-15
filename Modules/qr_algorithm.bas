Attribute VB_Name = "qr_algorithm"
Option Explicit

Const VSPACE As Integer = 1 'No. of empty rows between used range and result.
Const HSPACE As Integer = 1 'No. of empty cols between op name and result.
Const MSPACE As Integer = 1 'No. of empty cols between result matrices.
Const OPNAME As String = "QR decomposition"

Public Sub QR(Optional matrix_range As Variant, Optional upper_left As Variant)
    'Call the function for QR-decomposing a matrix_range and draw the result in an appropriate sheet.
    Dim row_count As Long, col_count As Long
    Dim sheet As Worksheet, dump_range As Range
    Dim a As Variant, q As Variant, r As Variant
    
    Set sheet = ActiveSheet
    
    If IsMissing(matrix_range) Then
        If TypeName(Selection) = "Range" Then
            Set matrix_range = Selection
        Else
            MsgBox TypeName(Selection) & " is not Range"
            Exit Sub
        End If
    End If
    If IsMissing(upper_left) Then
        Set upper_left = sheet.Cells(sheet.UsedRange.Rows.Count + 1 + VSPACE, 1)
    End If
    
    'Dimensions of the matrix_range.
    row_count = matrix_range.Rows.Count
    col_count = matrix_range.Columns.Count
    
    'Check whether m >= n where matrix_range in R^(m x n).
    If row_count < col_count Then
        upper_left.Cells(1, 1 + HSPACE).Value = "Number of columns exceeds the number of rows."
        Exit Sub
    End If
    
    'Carry out the computations.
    a = householder_qr(matrix_range.Value)
    
    'Transfer values.
    upper_left.Value = OPNAME
    
    'Draw Q.
    Set dump_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address, upper_left.Cells(row_count, row_count + HSPACE).Address)
    dump_range.Value = back_accumulation(a)
    'Draw R.
    Set dump_range = dump_range.Cells(1, row_count + 1 + MSPACE)
    Set dump_range = Range(dump_range.Address, dump_range.Cells(row_count, col_count).Address)
    dump_range.Value = matrixR(a)
End Sub

Function house(x As Variant) As Variant
    'Given a column vector x in R^m, this function computes a column vector v in R^m
    'with v(1) = 1 and beta in R such that P = I_m - beta * v @ v.T is orthogonal
    'and Px = ||x||_2 * e_1.
    Dim m0 As Integer, m As Integer, i As Integer, j As Integer
    Dim sigma As Variant, x1 As Variant, beta As Variant, mu As Variant
    
    m0 = LBound(x, 1): m = UBound(x, 1): j = LBound(x, 2)
    sigma = sum_of_squares(x, nstart:=m0 + 1)
    x1 = x(m0, j)
    x(m0, j) = 1 'x should be considered as v
    
    If sigma = 0 And x1 >= 0 Then
        beta = 0
    ElseIf sigma = 0 And x1 < 0 Then
        beta = -2
    Else
        mu = Sqr(x1 * x1 + sigma)
        If x1 <= 0 Then
            x(m0, j) = x1 - mu
        Else
            x(m0, j) = -sigma / (x1 + mu)
        End If
        beta = 2 * x(m0, j) ^ 2 / (sigma + x(m0, j) ^ 2)
        
        'v = v / v(1) but remember that x is v
        x1 = x(m0, j)
        For i = m0 To m
            x(i, j) = x(i, j) / x1
        Next i
    End If
    house = Array(x, beta)
End Function

Function householder_qr(a As Variant) As Variant
    'Given A in R^(m x n) with m >= n, the following algorithm
    'finds Householder matrices H_1, ..., H_n such that if Q = H_1 @ ... @ H_n,
    'then Q.T @ A = R is upper triangular. The upper triangular part of A is
    'overwritten by the upper triangular part of R and components j+1:m of the
    'jth Householder vector are stored in A(j+1:m,j), j < m.
    Dim m0 As Integer, m As Integer, n0 As Integer, n As Integer
    Dim j As Integer, k As Integer, kk As Integer
    Dim x() As Variant, v_and_beta As Variant
    Dim dummy As Variant
    
    m0 = LBound(a, 1): m = UBound(a, 1): n0 = LBound(a, 2): n = UBound(a, 2)
    
    For j = n0 To n
        ReDim x(j To m, j To j)
        'Populate x
        For k = j To m
            x(k, j) = a(k, j)
        Next k
        v_and_beta = house(x)
        
        'Carry out the update: A(j:m,j:n) = (I - beta * v @ v.T) @ A(j:m,j:n)
        dummy = dot(subtract_two(gen_eye(j, m, j, m), scalar_times_matrix(v_and_beta(1), dot(v_and_beta(0), mtranspose(v_and_beta(0))))), submatrix(a, j, m, j, n))
        For k = j To m
            For kk = j To n
                a(k, kk) = dummy(k, kk)
            Next kk
        Next k
        If j < m Then
            For k = j + 1 To m
                a(k, j) = v_and_beta(0)(k, j)
            Next k
        End If
    Next j
    householder_qr = a
End Function

Function back_accumulation(a As Variant, Optional k As Variant) As Variant
    Dim m0 As Integer, m As Integer, n0 As Integer, n As Integer
    Dim j As Integer, i As Integer, ii As Integer
    Dim q As Variant, v() As Variant, betaj As Variant, dummy As Variant
    
    m0 = LBound(a, 1): m = UBound(a, 1): n0 = LBound(a, 2): n = UBound(a, 2)
    If IsMissing(k) Then k = m
    
    q = gen_eye(m0, m, n0, k)
    For j = n To 1 Step -1
        'Populate v: v(j:m) = [1, A(j+1:m, j)].T
        ReDim v(j To m, j To j)
        v(j, j) = 1
        For i = j + 1 To m
            v(i, j) = a(i, j)
        Next i
        
        betaj = 2 / sum_of_squares(v)
        
        'Carry out the update: q(j:m, j:k) = q(j:m,j:k) - (betaj * v(j:m)) @ (v(j:m).T @ q(j:m,j:k))
        dummy = subtract_two(submatrix(q, j, m, j, k), dot(scalar_times_matrix(betaj, v), dot(mtranspose(v), submatrix(q, j, m, j, k))))
        For i = j To m
            For ii = j To k
                q(i, ii) = dummy(i, ii)
            Next ii
        Next i
    Next j
    back_accumulation = q
End Function

Function matrixR(a As Variant) As Variant
    'Recover R from a.
    Dim m0 As Integer, m As Integer, n0 As Integer, n As Integer
    Dim r() As Variant
    Dim i As Integer, j As Integer
    
    m0 = LBound(a, 1): m = UBound(a, 1): n0 = LBound(a, 2): n = UBound(a, 2)
    
    ReDim r(m0 To m, n0 To n)
    For i = 0 To m - m0
        For j = 0 To n - n0
            If i <= j Then
                r(m0 + i, n0 + j) = a(m0 + i, n0 + j)
            Else
                r(m0 + i, n0 + j) = 0
            End If
        Next j
    Next i
    matrixR = r
End Function
