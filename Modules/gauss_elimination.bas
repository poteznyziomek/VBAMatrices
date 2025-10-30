Attribute VB_Name = "gauss_elimination"
Option Explicit

'Implements gaussian elimination algorithms.

Function gauss_basic(mat_A As Variant, vec_b As Variant) As Variant
    'OBSOLETE.
    'Basic algorithm (without pivoting).
    Dim L As Variant
    Dim k As Integer, i As Integer, j As Integer, n As Integer
    n = UBound(mat_A, 1)
    
    For k = 1 To n - 1
        For i = k + 1 To n
            L(i, k) = mat_A(i, k) / mat_A(k, k)
            vec_b(i) = vec_b(i) - L(i, k) * vec_b(k)
            For j = k + 0 To n
                mat_A(i, j) = mat_A(i, j) - L(i, k) * mat_A(k, j)
            Next j
        Next i
    Next k
    gauss_basic = Array(L, mat_A, b)
End Function

Function gauss(mat_A As Variant, vec_b As Variant) As Variant
    'OBSOLETE.
    'Return a matrix and a vector after applying gaussian elimination with partial pivoting.
    Dim L As Variant, p As Variant, kcolumn() As Variant, swap_count As Integer, pmatrix, temp As Variant
    Dim k As Integer, i As Integer, j As Integer, n As Integer, ii As Integer, ind_max As Integer
    n = UBound(mat_A, 1)
    
    L = gen_eye(LBound(mat_A, 1), n, LBound(mat_A, 2), n)
    p = range_gen( _
        start_num:=1, _
        n:=UBound(mat_A, 1), _
        step:=1, _
        start_ind:=1)
    swap_count = 0
    
    'Algorithm
    '2.
    For k = 1 To n - 1
        'Populate column with n-k+1 abs of elements below (inclusive) the k-th column.
        '(index, value) in kcolumn.
        ReDim kcolumn(k To n, 1 To 2)
        For ii = k To n
            kcolumn(ii, 1) = ii
            kcolumn(ii, 2) = Abs(mat_A(ii, k))
        Next ii
        'Find the index of the maximum value (wrt abs) of the k-th column.
        ind_max = argmax(kcolumn)

        If ind_max <> k Then swap_count = swap_count + 1
        
        'Swap rows
        p = swap(p, ind_max, k)
        mat_A = mswap(mat_A, ind_max, k)
        vec_b = swap(vec_b, ind_max, k)
        If k > LBound(mat_A, 2) Then
            For ii = LBound(mat_A, 2) To k - 1
                temp = L(k, ii)
                L(k, ii) = L(ind_max, ii)
                L(ind_max, ii) = temp
            Next ii
        End If
        
        'Check whether mat_A(k,k) is zero.
        If mat_A(k, k) = 0 Then
            Debug.Print "No unique solution exists."
            gauss = -1
            Exit Function
        End If
        
        'a)
        For i = k + 1 To n
            'i)
            L(i, k) = mat_A(i, k) / mat_A(k, k)
            
            'ii)
            vec_b(i) = vec_b(i) - L(i, k) * vec_b(k)
            
            'iii)
            For j = k + 0 To n
                mat_A(i, j) = mat_A(i, j) - L(i, k) * mat_A(k, j)
            Next j
        Next i
    Next k
    
    gauss = Array(mat_A, vec_b)
End Function

Public Function gauss_pp(mat_A As Variant) As Variant
    'Return a matrix containing an LU decomposition of mat_A and permutation vector p.
    Dim i As Integer, j As Integer, k As Integer, n0 As Variant, n As Variant
    Dim p() As Variant
    Dim sign As Variant
    Dim copy As Variant, maxw As Variant, maxe As Variant, akk As Variant
    
    n0 = LBound(mat_A, 1)
    n = UBound(mat_A, 1)
    sign = 1 '1 if even number of row swaps else -1
    ReDim p(n0 To n)
    For i = n0 To n
        p(i) = i
    Next i
    
    For k = n0 To n - 1 '!!!!!!!!!!!!!!!!!!!
        maxw = k
        
        maxe = Abs(mat_A(p(k), k))
        
        For i = k + 1 To n
            If Abs(mat_A(p(i), k)) > maxe Then
                maxw = i
                maxe = Abs(mat_A(p(i), k))
            End If
        Next i

        If maxw <> p(k) Then
            sign = -sign
            p = swap(p, k, maxw)
        End If
        
        akk = mat_A(p(k), k)
        If akk <> 0 Then
        
            For i = k + 1 To n
                mat_A(p(i), k) = mat_A(p(i), k) / akk
            Next i
            
            For i = k + 1 To n
                For j = k + 1 To n
                    mat_A(p(i), j) = mat_A(p(i), j) - mat_A(p(i), k) * mat_A(p(k), j)
                Next j
            Next i
        End If
    Next k
    gauss_pp = Array(mat_A, p, sign)
    Debug.Print "end sign "; sign
End Function

Function pinv(p As Variant) As Variant
    'For a permutation p return its inverse.
    Dim s() As Variant
    Dim i As Integer
    ReDim s(LBound(p) To UBound(p))
    
    For i = LBound(p) To UBound(p)
        s(p(i)) = i
    Next i
    
    pinv = s
End Function
