Attribute VB_Name = "lin_sys_exact"
Option Explicit

Const VSPACE As Integer = 2 'No. of empty rows between used range and result.
Const HSPACE As Integer = 1 'No. of empty cols between op name and result.
Const OPNAME As String = "linear system solution"

Public Sub linear_system_solve(Optional matrix_range As Variant, Optional upper_left As Variant)
    'Call the function for solving a system of linear equations and draw the result in an appropriate sheet.
    Dim row_count As Long, col_count As Long
    Dim system_conditions As Boolean
    Dim exes() As String
    Dim gauss_result As Variant
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
    
    Dim i As Integer, b As Variant
'    ReDim b(1 To 4)
'    b = matrix_range.Areas(2).Value
'    b = Range("f2:f5").Value
'    ReDim b(1 To 4)
'    b = Application.Transpose(matrix_range.Areas(2).Value)
'    Debug.Print LBound(b); " "; UBound(b)
'    For i = 1 To 4
'        Debug.Print b(i)
'    Next i
'
    'Dump the result in the sheet.
    row_count = matrix_range.Rows.Count
    col_count = matrix_range.Columns.Count
'    Set dump_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address, upper_left.Cells(row_count, col_count + HSPACE).Address)
    
    upper_left.Value = OPNAME
    gauss_result = gauss(matrix_range.Areas(1).Value, _
        Application.Transpose(matrix_range.Areas(2).Value))
    ReDim exes(1 To row_count)
    For i = 1 To row_count
        exes(i) = "x" & i
    Next i
    Debug.Print TypeName(gauss_result)
    If TypeName(gauss_result) = "Integer" Then
        Set dump_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address)
        dump_range.Value = "No unique solutions."
    Else
        Set dump_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address, upper_left.Cells(row_count, 1 + HSPACE).Address)
        dump_range.Value = Application.Transpose(exes)
        Set dump_range = sheet.Range(upper_left.Cells(1, 2 + HSPACE).Address, upper_left.Cells(row_count, 2 + HSPACE).Address)
        dump_range.Value = Application.Transpose(u_solve(gauss_result(0), gauss_result(1)))
    End If
End Sub

Sub call_lin_sys_solve()
    Call linear_system_solve
End Sub

Function u_solve(u As Variant, b As Variant) As Variant
    'Return the solution to the system of linear equations with an upper triangular square matrix u.
    Dim x() As Variant
    Dim n0 As Integer, n As Integer, i As Integer, j As Integer, k As Integer, kk As Integer
    Dim m0 As Integer, m As Integer
    
    n0 = LBound(u, 1): n = UBound(u, 1)
    m0 = LBound(u, 2): m = UBound(u, 2)
    
    If n - n0 + 1 <> m - m0 + 1 Then
        Debug.Print "Not a square matrix."
        u_solve = -1
        Exit Function
    End If
    
    ReDim x(1 To n - n0 + 1)
    For i = n To 1 Step -1
        x(i) = 0
        For j = i + 1 To n
            x(i) = x(i) + u(i, j) * x(j)
        Next j
        x(i) = (b(i) - x(i)) / u(i, i)
    Next i
    u_solve = x
End Function

Sub test_u_solve()
    Dim u() As Variant, b() As Variant, i As Integer
    ReDim u(1 To 4, 1 To 4): ReDim b(1 To 4)
    
'    u(1, 1) = 2: u(1, 2) = 3: u(1, 3) = -4: b(1) = 1
'    u(2, 1) = 0: u(2, 2) = 7: u(2, 3) = 10: b(2) = 17
'    u(3, 1) = 0: u(3, 2) = 0: u(3, 3) = 4: b(3) = 4
    
'    u(1, 1) = 1: u(1, 2) = 1: u(1, 3) = -2: u(1, 4) = 1: b(1) = 1
'    u(2, 1) = 0: u(2, 2) = 1: u(2, 3) = 5: u(2, 4) = -5: b(2) = 1
'    u(3, 1) = 0: u(3, 2) = 0: u(3, 3) = 8: u(3, 4) = -8: b(3) = 0
'    u(4, 1) = 0: u(4, 2) = 0: u(4, 3) = 0: u(4, 4) = 4: b(4) = 4
    
    u(1, 1) = 1: u(1, 2) = -1: u(1, 3) = 2: u(1, 4) = -1: b(1) = -8
    u(2, 1) = 0: u(2, 2) = 2: u(2, 3) = -1: u(2, 4) = 1: b(2) = 6
    u(3, 1) = 0: u(3, 2) = 0: u(3, 3) = -1: u(3, 4) = -1: b(3) = -4
    u(4, 1) = 0: u(4, 2) = 0: u(4, 3) = 0: u(4, 4) = 2: b(4) = 4
    
    For i = 1 To 4
        Debug.Print u_solve(u, b)(i)
    Next
End Sub
