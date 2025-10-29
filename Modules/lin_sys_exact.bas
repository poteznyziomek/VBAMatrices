Attribute VB_Name = "lin_sys_exact"
Option Explicit

Const VSPACE As Integer = 2 'No. of empty rows between used range and result.
Const HSPACE As Integer = 1 'No. of empty cols between op name and result.
Const OPNAME As String = "linear system solution"

Public Sub linear_system_solve(Optional matrix_range As Variant, Optional upper_left As Variant)
    'Call the function for solving a system of linear equations and draw the result in an appropriate sheet.
    Dim row_count As Long, col_count As Long, i As Integer
    Dim system_conditions As Boolean
    Dim exes() As String
    Dim gauss_result As Variant, LU As Variant
    Dim dump_range As Range
    Dim sheet As Worksheet
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
    
    'Check whether there is a matrix and a vector and they are compatible.
    system_conditions = matrix_range.Areas.Count <> 2 _
        Or matrix_range.Areas(1).Rows.Count <> matrix_range.Areas(1).Columns.Count _
        Or matrix_range.Areas(2).Columns.Count <> 1 _
        Or matrix_range.Areas(1).Rows.Count <> matrix_range.Areas(2).Rows.Count
    If system_conditions Then
        upper_left.Value = OPNAME
        upper_left.Cells(1, 1 + HSPACE).Value = "Invalid input."
        Exit Sub
    End If
    
    'Dump the result in the sheet.
    row_count = matrix_range.Rows.Count
    col_count = matrix_range.Columns.Count
'    Set dump_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address, upper_left.Cells(row_count, col_count + HSPACE).Address)
    
    upper_left.Value = OPNAME
    LU = gauss_pp(matrix_range.Areas(1).Value) 'Holds a matrix containing both L and U, and a permutation vector p if input is invertible, else ...
'    gauss_result = gauss(matrix_range.Areas(1).Value, _
'        Application.Transpose(matrix_range.Areas(2).Value))
'    gauss_result = lu_solve(LU, Application.Transpose(matrix_range.Areas(2).Value))
    ReDim exes(1 To row_count)
    For i = 1 To row_count
        exes(i) = "x" & i
    Next i
'    If TypeName(gauss_result) = "Integer" Then
    If TypeName(LU) = "Integer" Then
        Set dump_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address)
        dump_range.Value = "No unique solutions."
    Else
        Set dump_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address, upper_left.Cells(row_count, 1 + HSPACE).Address)
        dump_range.Value = Application.Transpose(exes)
        Set dump_range = sheet.Range(upper_left.Cells(1, 2 + HSPACE).Address, upper_left.Cells(row_count, 2 + HSPACE).Address)
'        dump_range.Value = Application.Transpose(u_solve(gauss_result(0), gauss_result(1)))
        gauss_result = lu_solve(LU, Application.Transpose(matrix_range.Areas(2).Value))
        dump_range.Value = Application.Transpose(gauss_result)
    End If
End Sub

Sub call_lin_sys_solve()
    Call linear_system_solve
End Sub

Function u_solve(U As Variant, b As Variant) As Variant
    'Return the solution to the system of linear equations with an upper triangular square matrix u.
    Dim x() As Variant
    Dim n0 As Integer, n As Integer, i As Integer, j As Integer, k As Integer, kk As Integer
    Dim m0 As Integer, m As Integer
    
    n0 = LBound(U, 1): n = UBound(U, 1)
    m0 = LBound(U, 2): m = UBound(U, 2)
    
    If n - n0 + 1 <> m - m0 + 1 Then
        Debug.Print "Not a square matrix."
        u_solve = -1
        Exit Function
    End If
    
    ReDim x(1 To n - n0 + 1)
    For i = n To 1 Step -1
        x(i) = 0
        For j = i + 1 To n
            x(i) = x(i) + U(i, j) * x(j)
        Next j
        x(i) = (b(i) - x(i)) / U(i, i)
    Next i
    u_solve = x
End Function

Function lu_solve(LU As Variant, b As Variant) As Variant
    'Return the solution to the system Ax = b where A satisfies PA = LU.
    Dim x() As Variant, y() As Variant, p As Variant, pinv As Variant
    Dim n0 As Integer, n As Integer, i As Integer, j As Integer
    p = LU(1)
    pinv = permutation_inverse(p)
    n0 = LBound(LU(1), 1)
    n = UBound(LU(1), 1)
    ReDim x(n0 To n)
    ReDim y(n0 To n)
    
    'Forward substitution.
    For i = n0 To n
        y(i) = 0
        For j = n0 To i - 1
            y(i) = y(i) + LU(0)(p(i), j) * y(j)
        Next j
'        Debug.Print p(i)
'        y(i) = (b(p(i)) - y(i)) / LU(0)(p(i), i)
        y(i) = b(p(i)) - y(i)
    Next i
    'Backward substitution.
    For i = n To n0 Step -1
        x(i) = 0
        For j = i + 1 To n
            x(i) = x(i) + LU(0)(p(i), j) * x(j)
        Next j
        x(i) = (y(i) - x(i)) / LU(0)(p(i), i)
    Next i
    lu_solve = x
End Function

Sub test_u_solve()
    Dim U() As Variant, b() As Variant, i As Integer, res As Variant, LU As Variant
    ReDim U(1 To 4, 1 To 4): ReDim b(1 To 4)
    
'    u(1, 1) = 2: u(1, 2) = 3: u(1, 3) = -4: b(1) = 1
'    u(2, 1) = 0: u(2, 2) = 7: u(2, 3) = 10: b(2) = 17
'    u(3, 1) = 0: u(3, 2) = 0: u(3, 3) = 4: b(3) = 4
    
'    u(1, 1) = 1: u(1, 2) = 1: u(1, 3) = -2: u(1, 4) = 1: b(1) = 1
'    u(2, 1) = 0: u(2, 2) = 1: u(2, 3) = 5: u(2, 4) = -5: b(2) = 1
'    u(3, 1) = 0: u(3, 2) = 0: u(3, 3) = 8: u(3, 4) = -8: b(3) = 0
'    u(4, 1) = 0: u(4, 2) = 0: u(4, 3) = 0: u(4, 4) = 4: b(4) = 4
    
'    U(1, 1) = 1: U(1, 2) = -1: U(1, 3) = 2: U(1, 4) = -1: b(1) = -8
'    U(2, 1) = 0: U(2, 2) = 2: U(2, 3) = -1: U(2, 4) = 1: b(2) = 6
'    U(3, 1) = 0: U(3, 2) = 0: U(3, 3) = -1: U(3, 4) = -1: b(3) = -4
'    U(4, 1) = 0: U(4, 2) = 0: U(4, 3) = 0: U(4, 4) = 2: b(4) = 4

    U(1, 1) = 2: U(1, 2) = 0: U(1, 3) = 0: U(1, 4) = 0: b(1) = 3
    U(2, 1) = 1: U(2, 2) = 1.5: U(2, 3) = 0: U(2, 4) = 0: b(2) = 4.5
    U(3, 1) = 0: U(3, 2) = -3: U(3, 3) = 0.5: U(3, 4) = 0: b(3) = -6.6
    U(4, 1) = 2: U(4, 2) = -2: U(4, 3) = 1: U(4, 4) = 1: b(4) = 0.8
    
    LU = gauss_pp(U)
    
    res = lu_solve(LU, b)
    For i = 1 To 4
        Debug.Print res(i)
    Next i
End Sub
