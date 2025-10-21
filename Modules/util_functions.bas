Attribute VB_Name = "util_functions"
Option Explicit

Const RELTOLDEFAULT = 0.000000001 'Default parameter for isclose function.
Const ABSTOLDEFAULT = 0#  'Default parameter for isclose function.

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

Public Function arr_len(arr As Variant) As Variant
    'Return the number of elements of arr.
    arr_len = UBound(arr) - LBound(arr) + 1
End Function

Public Function count_report_sheets(sought_name As String) As Integer
    'Return the number of sheets whose name starts with sought_name.
    Dim sheet As Worksheet
    Dim book As Workbook
    Set book = ActiveWorkbook

    count_report_sheets = 0
    For Each sheet In book.Worksheets
        If VBA.Strings.InStr(Start:=1, String1:=sheet.name, String2:=sought_name, Compare:=vbTextCompare) > 0 _
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
            If VBA.Strings.InStr(Start:=1, String1:=sheet.name, String2:=preff, Compare:=vbTextCompare) > 0 _
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

Function isclose(a As Double, b As Double, Optional rel_tol, Optional abs_tol) As Boolean
    'Return True if a and b are close to each other and False otherwise.
    Dim numeric_condition As Boolean, domain_condition As Boolean
    
'    numeric_condition = IsNumeric(a) And IsNumeric(b) _
'        And IsNumeric(rel_tol) And IsNumeric(abs_tol)
'    Debug.Print TypeName(a); TypeName(b)
'    Debug.Print numeric_condition
'    If Not numeric_condition Then
'        Debug.Print "Non-numeric argument(s) passed."
'    End If
    
    If IsMissing(rel_tol) Then rel_tol = RELTOLDEFAULT
    If IsMissing(abs_tol) Then abs_tol = ABSTOLDEFAULT
    
    domain_condition = 0# <= rel_tol And rel_tol < 1# And abs_tol >= 0#
    If Not domain_condition Then
        Debug.Print "Invalid tolerance(s)."
    End If
    
    isclose = Abs(a - b) <= maxi(rel_tol * maxi(Abs(a), Abs(b)), abs_tol)
End Function

Function swap(p As Variant, i As Integer, j As Integer) As Variant
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
    'Swap i-th and j-th row in m.
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

Function floor(x As Double) As Variant
    'Return the floor(x).
    floor = Int(x) - 1 * (Int(x) > x)
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

Sub test_isclose()
    Dim triples() As Double, i As Integer
    ReDim triples(1 To 14, 1 To 4)
    
    'a                              b                           rel_tol                     abs_tol
    triples(1, 1) = 10000000000#: triples(1, 2) = 10000100000#: triples(1, 3) = 0.00001: triples(1, 4) = 0.00000001
    triples(2, 1) = 0.0000001: triples(2, 2) = 0.00000001: triples(2, 3) = 0.00001: triples(2, 4) = 0.00000001
    triples(3, 1) = 10000000000#: triples(3, 2) = 10000100000#: triples(3, 3) = 0.00001: triples(3, 4) = 0.00000001
    triples(4, 1) = 0.00000001: triples(4, 2) = 0.000000001: triples(4, 3) = 0.00001: triples(4, 4) = 0.00000001
    triples(5, 1) = 10000000000#: triples(5, 2) = 10001000000#: triples(5, 3) = 0.00001: triples(5, 4) = 0.00000001
    triples(6, 1) = 0.00000001: triples(6, 2) = 0.000000001: triples(6, 3) = 0.00001: triples(6, 4) = 0.00000001
    triples(7, 1) = 0.00000001: triples(7, 2) = 0#: triples(7, 3) = 0.00001: triples(7, 4) = 0.00000001
    triples(8, 1) = 0.0000001: triples(8, 2) = 0#: triples(8, 3) = 0.00001: triples(8, 4) = 0.00000001
    triples(9, 1) = 1E-100: triples(9, 2) = 0#: triples(9, 3) = 0.00001: triples(9, 4) = 0#
    triples(10, 1) = 0.0000001: triples(10, 2) = 0#: triples(10, 3) = 0.00001: triples(10, 4) = 0#
    triples(11, 1) = 0.0000000001: triples(11, 2) = 1E-20: triples(11, 3) = 0.00001: triples(11, 4) = 0.00000001
    triples(12, 1) = 0.0000000001: triples(12, 2) = 0#: triples(12, 3) = 0.00001: triples(12, 4) = 0.00000001
    triples(13, 1) = 0.0000000001: triples(13, 2) = 1E-20: triples(13, 3) = 0.00001: triples(13, 4) = 0#
    triples(14, 1) = 0.0000000001: triples(14, 2) = 9.99999E-11: triples(14, 3) = 0.00001: triples(14, 4) = 0#
    
    For i = 1 To 14
        Debug.Print i & ": a = " & triples(i, 1) & ", b = " & triples(i, 2); ", " & isclose(a:=triples(i, 1), b:=triples(i, 2), rel_tol:=triples(i, 3))
    Next i
End Sub
