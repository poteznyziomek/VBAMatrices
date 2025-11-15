Attribute VB_Name = "util_functions"
Option Explicit

Const RELTOLDEFAULT = 0.000000001 'Default parameter for isclose function.
Const ABSTOLDEFAULT = 0#  'Default parameter for isclose function.

' || Numeric functions.
Public Function mini(x As Variant, y As Variant) As Variant
    'Return the minimum of x and y.
    If x > y Then
        mini = y
    Else
        mini = x
    End If
End Function

Public Function maxi(x As Variant, y As Variant) As Variant
    'Return the maximum of x and y.
    If x > y Then
        maxi = x
    Else
        maxi = y
    End If
End Function

Public Function maxi_n(arr As Variant) As Variant
    'Return the maximum of the elements of arr.
    Dim i As Long
    
    maxi_n = arr(LBound(arr))
    For i = LBound(arr) + 1 To UBound(arr)
        If arr(i) > maxi_n Then
            maxi_n = arr(i)
        End If
    Next i
End Function

Function isclose(a As Double, b As Double, Optional rel_tol, Optional abs_tol) As Boolean
    'Return True if a and b are close to each other and False otherwise.
    Dim numeric_condition As Boolean, domain_condition As Boolean
    
    If IsMissing(rel_tol) Then rel_tol = RELTOLDEFAULT
    If IsMissing(abs_tol) Then abs_tol = ABSTOLDEFAULT
    
    domain_condition = 0# <= rel_tol And rel_tol < 1# And abs_tol >= 0#
    If Not domain_condition Then
        Debug.Print "Invalid tolerance(s)."
    End If
    
    isclose = Abs(a - b) <= maxi(rel_tol * maxi(Abs(a), Abs(b)), abs_tol)
End Function

Function floor(x As Double) As Variant
    'Return the floor(x).
    floor = Int(x) - 1 * (Int(x) > x)
End Function

' || Sheet functions.
Public Function count_report_sheets(sought_name As String) As Integer
    'Return the number of sheets whose name starts with sought_name.
    Dim sheet As Worksheet
    Dim book As Workbook
    Set book = ActiveWorkbook

    count_report_sheets = 0
    For Each sheet In book.Worksheets
        If VBA.Strings.InStr(start:=1, String1:=sheet.name, String2:=sought_name, Compare:=vbTextCompare) > 0 _
          And IsNumeric(Right(sheet.name, 1)) Then
              count_report_sheets = count_report_sheets + 1
        End If
    Next sheet
End Function

Public Sub placeholder_sub(matrix_range As Variant, upper_left As Variant)
    'Placeholder for procedures not yet implemented.
    upper_left.Value = "Method not yet implemented."
    Debug.Print TypeName(upper_left)
End Sub

Public Function create_report_sheet_name(preff As String) As String
    'Return a string of the form "pref N" where N is a positive integer.
    Dim no_reports As Integer, i As Integer
    Dim report_numbers() As Integer
    Dim new_report_name As String, suff As String
    Dim sheet As Worksheet
    Dim book As Workbook
    Set book = ActiveWorkbook
    
    no_reports = count_report_sheets(preff)
    
    If no_reports > 0 Then
        ReDim report_numbers(1 To no_reports)
        i = 1
        For Each sheet In book.Worksheets
            suff = Mid(sheet.name, Len(preff) + 1)
            If VBA.Strings.InStr(start:=1, String1:=sheet.name, String2:=preff, Compare:=vbTextCompare) > 0 _
              And IsNumeric(suff) Then
                  report_numbers(i) = CInt(suff)
                  i = i + 1
            End If
        Next sheet
        create_report_sheet_name = "Report " & CStr(maxi_n(report_numbers) + 1)
    Else
        create_report_sheet_name = "Report 1"
    End If
End Function

' || Matrix/vector functions.
Function eye(n As Long) As Variant
    'Return the identity matrix of degree n.
    Dim return_arr() As Variant, i As Long, j As Long
    ReDim return_arr(1 To n, 1 To n)
    
    For i = 1 To n
        For j = 1 To n
            If i = j Then
                return_arr(i, j) = 1
            Else
                return_arr(i, j) = 0
            End If
        Next j
    Next i
    eye = return_arr
End Function

Function gen_eye(nstart As Variant, nend As Variant, mstart As Variant, mend As Variant) As Variant
    'Return the matrix with 1s on the main diagonal.
    Dim return_arr() As Variant, i As Long, j As Long
    ReDim return_arr(nstart To nend, mstart To mend)
    
    For i = 0 To nend - nstart
        For j = 0 To mend - mstart
            If i = j Then
                return_arr(nstart + i, mstart + j) = 1#
            Else
                return_arr(nstart + i, mstart + j) = 0#
            End If
        Next j
    Next i
    gen_eye = return_arr
End Function

Function permutation_matrix(p As Variant) As Variant
    'Return a permutation matrix based on the permutation p.
    Dim lb As Integer, ub As Integer, i As Integer, j As Integer
    Dim pmatrix() As Integer
    lb = LBound(p): ub = UBound(p)
    
    ReDim pmatrix(lb To ub, lb To ub)
    For i = lb To ub
        For j = lb To ub
            If p(i) = j Then
                pmatrix(i, j) = 1
            Else
                pmatrix(i, j) = 0
            End If
        Next j
    Next i
    permutation_matrix = pmatrix
End Function

Function permutation_inverse(p As Variant) As Variant
    'Return the inverse of a permutation p.
    Dim pinv() As Variant, i As Integer
    ReDim pinv(LBound(p) To UBound(p))
    For i = LBound(p) To UBound(p)
        pinv(p(i)) = i
    Next i
    permutation_inverse = pinv
End Function

Function range_gen(start_num As Variant, n As Integer, Optional step = 1, Optional start_ind = 0) As Variant
    'Return an n-element array with elements
    'arr(k) = start_num + i * step, where start_ind <= k <= n + start_ind - 1 and 0 <= i <= n - 1.
    
    Dim i As Integer, k As Integer
    Dim arr() As Variant
    
    ReDim arr(start_ind To n + start_ind - 1)
    
    k = start_ind
    For i = 0 To n - 1
        arr(k) = start_num + i * step
        k = k + 1
    Next i
    
    range_gen = arr
End Function

Function swap(p As Variant, i As Variant, j As Variant) As Variant
    'Swap i-th and j-th element in p.
    Dim temp As Variant, range_condition As Boolean
    
    range_condition = LBound(p) <= i And i <= UBound(p) _
        And LBound(p) <= j And j <= UBound(p)
    If Not range_condition Then
        Debug.Print "Indices out of bounds."
        swap = -1
        Exit Function
    End If
    
    If i = j Then
        swap = p
    Else
        temp = p(i)
        p(i) = p(j)
        p(j) = temp
        swap = p
    End If
End Function

Function mswap(m As Variant, i As Integer, j As Integer) As Variant
    'Swap i-th and j-th row in matrix m.
    Dim temp() As Variant, range_condition As Boolean, k As Integer
    
    range_condition = LBound(m, 1) <= i And i <= UBound(m, 1) _
        And LBound(m, 2) <= j And j <= UBound(m, 2)
    If Not range_condition Then
        Debug.Print "Indices out of bound."
        mswap = -1
        Exit Function
    End If
    
    If i = j Then
        mswap = m
    Else
        ReDim temp(LBound(m, 2) To UBound(m, 2))
        For k = LBound(m, 2) To UBound(m, 2)
            temp(k) = m(i, k)
            m(i, k) = m(j, k)
            m(j, k) = temp(k)
        Next k
        mswap = m
    End If
End Function

Public Function arr_len(arr As Variant) As Variant
    'Return the number of elements of arr.
    arr_len = UBound(arr) - LBound(arr) + 1
End Function

Function argmax(arr As Variant) As Variant
    'Return the index of the maximum element of the arr.
    Dim i As Long, biggest As Variant
    ReDim biggest(LBound(arr, 1) To UBound(arr, 1))
    For i = LBound(arr, 1) To UBound(arr, 1)
        biggest(i) = arr(i, 2)
    Next i
    biggest = maxi_n(biggest)
    argmax = arr(LBound(arr, 1), LBound(arr, 2))
    
    For i = LBound(arr, 1) + 0 To UBound(arr, 1)
        If arr(i, 2) = biggest Then
            argmax = arr(i, 1)
            Exit Function
        End If
    Next i
End Function

Function mtranspose(mat_A As Variant) As Variant
    'Return the transpose of the matrix mat_A.
    Dim n0 As Integer, n As Integer
    Dim m0 As Integer, m As Integer
    Dim i As Integer, j As Integer
    Dim trans() As Variant
    
    n0 = LBound(mat_A, 1): n = UBound(mat_A, 1)
    m0 = LBound(mat_A, 2): m = UBound(mat_A, 2)
    ReDim trans(m0 To m, n0 To n)
    For i = n0 To n
        For j = m0 To m
            trans(j, i) = mat_A(i, j)
        Next j
    Next i
    mtranspose = trans
End Function

Function sum_of_squares(v As Variant, Optional nstart As Variant, Optional nstop As Variant) As Variant
    'Return the partial sum of squares of a column vector v.
    Dim m0 As Integer, m As Integer, j As Integer, i As Integer
    
    m0 = LBound(v, 1): m = UBound(v, 1): j = LBound(v, 2)
    If IsMissing(nstart) Then nstart = m0
    If IsMissing(nstop) Then nstop = m

    sum_of_squares = 0
    For i = nstart To nstop
        sum_of_squares = sum_of_squares + v(i, j) ^ 2
    Next i
End Function

Function scalar_times_matrix(alpha As Variant, a As Variant) As Variant
    'Return a matrix.
    Dim m0 As Integer, m As Integer, n0 As Integer, n As Integer
    Dim i As Integer, j As Integer
    Dim alpha_m()
    
    m0 = LBound(a, 1): m = UBound(a, 1): n0 = LBound(a, 2): n = UBound(a, 2)
    ReDim alpha_m(m0 To m, n0 To n)
    
    For i = m0 To m
        For j = n0 To n
            alpha_m(i, j) = alpha * a(i, j)
        Next j
    Next i
    scalar_times_matrix = alpha_m
End Function

Function submatrix(a As Variant, Optional mstart As Variant, Optional mend As Variant, Optional nstart As Variant, Optional nend As Variant) As Variant
    'Return a submatrix a(mstart:mend, nstart:nend).
    Dim i As Integer, j As Integer
    Dim result() As Variant
    
    If IsMissing(mstart) Then mstart = LBound(a, 1)
    If IsMissing(mend) Then mend = UBound(a, 1)
    If IsMissing(nstart) Then nstart = LBound(a, 2)
    If IsMissing(nend) Then nend = UBound(a, 2)
    
    ReDim result(mstart To mend, nstart To nend)
    For i = mstart To mend
        For j = nstart To nend
            result(i, j) = a(i, j)
        Next j
    Next i
    submatrix = result
End Function
