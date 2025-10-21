Attribute VB_Name = "gauss_elimination"
Option Explicit

'Implements gaussian elimination algorithms.

Function gauss_basic(mat_A As Variant, vec_b As Variant) As Variant
    'Basic algorithm (without pivoting)
'    Dim l() As Double
    Dim l As Variant
    Dim k As Integer, i As Integer, j As Integer, n As Integer
    n = UBound(mat_A, 1)
'    ReDim l(LBound(mat_A, 1) To n, LBound(mat_A, 2) To n)
'    For i = LBound(l, 1) To n
'        For j = LBound(l, 2) To n
'            If i = j Then
'                l(i, j) = 1
'            Else
'                l(i, j) = 0
'            End If
'        Next j
'    Next i
'    l = gen_eye(LBound(mat_A, 1), UBound(mat_A, 1), LBound(mat_A, 2), UBound(mat_A, 2))
    
    'Algorithm.
    For k = 1 To n - 1
        For i = k + 1 To n
            l(i, k) = mat_A(i, k) / mat_A(k, k)
            vec_b(i) = vec_b(i) - l(i, k) * vec_b(k)
            For j = k + 0 To n
                mat_A(i, j) = mat_A(i, j) - l(i, k) * mat_A(k, j)
            Next j
        Next i
    Next k
    
    gauss_basic = l
End Function

Function gauss(mat_A As Variant, vec_b As Variant) As Variant
    'Return a matrix and a vector after applying gaussian elimination with partial pivoting.
    Dim l As Variant, p As Variant, kcolumn() As Variant, swap_count As Integer, pmatrix, temp As Variant
    Dim k As Integer, i As Integer, j As Integer, n As Integer, ii As Integer, ind_max As Integer
    n = UBound(mat_A, 1)
    
    l = gen_eye(LBound(mat_A, 1), n, LBound(mat_A, 2), n)
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
                temp = l(k, ii)
                l(k, ii) = l(ind_max, ii)
                l(ind_max, ii) = temp
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
            l(i, k) = mat_A(i, k) / mat_A(k, k)
            
            'ii)
            vec_b(i) = vec_b(i) - l(i, k) * vec_b(k)
            
            'iii)
            For j = k + 0 To n
                mat_A(i, j) = mat_A(i, j) - l(i, k) * mat_A(k, j)
            Next j
        Next i
    Next k
    
    gauss = Array(mat_A, vec_b)
End Function

Sub test_gauss_general()
    Dim n As Integer, i As Integer, j As Integer
    Dim mat_A() As Variant, vec_b() As Variant
    Dim result As Variant
    Dim dump_range As Range
    
    n = 4
    ReDim mat_A(1 To n, 1 To n): ReDim vec_b(1 To n)
    mat_A(1, 1) = 1: mat_A(1, 2) = 1: mat_A(1, 3) = -2: mat_A(1, 4) = 1: vec_b(1) = 1
    mat_A(2, 1) = 1: mat_A(2, 2) = 2: mat_A(2, 3) = 3: mat_A(2, 4) = -4: vec_b(2) = 2
    mat_A(3, 1) = 2: mat_A(3, 2) = 1: mat_A(3, 3) = -1: mat_A(3, 4) = -1: vec_b(3) = 1
    mat_A(4, 1) = 1: mat_A(4, 2) = -1: mat_A(4, 3) = 1: mat_A(4, 4) = 2: vec_b(4) = 3
    
    'result = gauss_basic(mat_A, vec_b)
    result = gauss(mat_A, vec_b)
    
    For i = 1 To n
        For j = 1 To n
'            Debug.Print result(i, j)
        Next j
    Next i
    
    Set dump_range = Selection
    dump_range.Value = result(2)
End Sub

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

Private Function gen_eye(nstart As Variant, nend As Variant, mstart As Variant, mend As Variant) As Variant
    'Return the matrix with 1s on the main diagonal.
    Dim return_arr() As Variant, i As Long, j As Long
    ReDim return_arr(nstart To nend, mstart To mend)
    
    For i = nstart To nend
        For j = mstart To mend
            If i = j Then
                return_arr(i, j) = 1#
            Else
                return_arr(i, j) = 0#
            End If
        Next j
    Next i
    gen_eye = return_arr
End Function
