Attribute VB_Name = "inverse"
Option Explicit

Const VSPACE As Integer = 2 'No. of empty rows between used range and result.
Const HSPACE As Integer = 1 'No. of empty cols between op name and result.
Const OPNAME = "Inverse"

Sub mat_inv(Optional matrix_range As Variant, Optional upper_left As Variant)
    'Call the matrix inverse function and draw the result in an appropriate sheet.
    Dim m_inverse As Variant
    Dim row_count As Long, col_count As Long
    Dim dump_range As Range, sheet As Worksheet
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
    
    'Check whether matrix is square.
    If matrix_range.Rows.Count <> matrix_range.Columns.Count Then
        upper_left.Value = OPNAME
        upper_left.Cells(1, 1 + HSPACE).Value = "Not a square matrix."
        Exit Sub
    End If
    
    m_inverse = inv(matrix_range.Value)
    If TypeName(m_inverse) = "Integer" Then
        Set dump_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address)
        upper_left.Value = OPNAME
        dump_range.Value = "Does not exist."
    Else
        row_count = matrix_range.Rows.Count
        col_count = matrix_range.Columns.Count
        Set dump_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address, upper_left.Cells(row_count, col_count + HSPACE).Address)
        upper_left.Value = OPNAME
        dump_range.Value = m_inverse
    End If
End Sub

Function inv(mat_A As Variant) As Variant
    'Return the inverse of the matrix mat_A if it exists, -1 otherwise.
    
    '1. LU-decompose mat_A.
    '2. In a loop (j = 1, 2, ..., n):
    '   a. solve Ly_j = Pe_j for y_j;
    '   b. solve Ux_j = y_j for x_j.
    '   c. x_j is the j-th column of mat_A^(-1).
    
    Dim m As Variant, p As Variant, pinv As Variant
    Dim n0 As Integer, n As Integer, j As Integer, i As Integer
    Dim yj As Variant
    Dim result() As Variant
    
    m = gauss_pp(mat_A)
    p = m(1) 'permutation vector of the decomposed mat_A
    m = m(0) 'holds both L and U of mat_A
    n0 = LBound(m, 1)
    n = UBound(m, 1)
    pinv = permutation_inverse(p)
    ReDim result(n0 To n, n0 To n)
    
    For j = n0 To n
        yj = forward_substitution(m, p, j)
        yj = backward_substitution(m, p, yj) 'Imagine that we're assigning to xj.
        'Fill the j-th column of result
        For i = n0 To n
            'Check wheter matrix is singular.
            If TypeName(yj) = "Integer" Then
                inv = -1
                Exit Function
            Else
                result(i, p(j)) = yj(i)
            End If
        Next i
    Next j
    inv = result
End Function

Function forward_substitution(m As Variant, p As Variant, j As Variant) As Variant
    'Solve the system L y_j = P e_j for y_j, where e_j has 1 at index j and zeros everywhere else.
    Dim n0 As Integer, n As Integer, i As Integer, k As Integer, nn As Integer
    Dim yj() As Variant
    
    nn = j
    n0 = LBound(m, 1)
    n = UBound(m, 1)
    ReDim yj(n0 To n)

    'Fill the first nn-1 entries of yj with zeros.
    For i = n0 To nn - 1
        yj(i) = 0
    Next i
    
    yj(nn) = 1
    
    'Fill the rest n-nn entries.
    For i = 1 To n - nn
        yj(nn + i) = 0
        For k = 1 To i
            yj(nn + i) = yj(nn + i) + m(p(nn + i), k + nn - 1) * yj(k + nn - 1)
        Next k
        yj(nn + i) = -yj(nn + i)
    Next i
    forward_substitution = yj
End Function

Function backward_substitution(m As Variant, p As Variant, yj As Variant) As Variant
    'Solve the system U x_j = y_j for x_j.
    Dim n0 As Integer, n As Integer, i As Integer, k As Integer
    Dim xj() As Variant
    
    n0 = LBound(m, 1)
    n = UBound(m, 1)
    ReDim xj(n0 To n)
    
    For i = n To n0 Step -1
        xj(i) = 0
        For k = i + 1 To n
            xj(i) = xj(i) + m(p(i), k) * xj(k)
        Next k
        
        'Check if matrix is singular
        If m(p(i), i) = 0 Then
            backward_substitution = -1
            Exit Function
        Else
            xj(i) = (yj(i) - xj(i)) / m(p(i), i)
        End If
    Next i
    backward_substitution = xj
End Function
