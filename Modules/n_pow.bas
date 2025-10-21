Attribute VB_Name = "n_pow"
Option Explicit

Const VSPACE As Integer = 1 'No. of empty rows between used range and result.
Const HSPACE As Integer = 1 'No. of empty cols between op name and result.
Const OPNAME As String = "Power N = "

Sub mat_npow(nth As Integer, Optional matrix_range As Variant, Optional upper_left As Variant)
    'Call a function to calculate nth power of a matrix and draw the result in an appropriate sheet.
    Dim i As Integer
    Dim dump_range As Range
    Dim sheet As Worksheet
    Dim row_count As Long, col_count As Long
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
    
    'Check whether there is one matrix.
    If matrix_range.Areas.Count > 1 Then
        upper_left.Value = OPNAME
        upper_left.Cells(1, 1 + HSPACE).Value = "Invalid number of matrices."
        Exit Sub
    End If
    
    'Is it a square matrix?
    row_count = matrix_range.Rows.Count
    col_count = matrix_range.Columns.Count
    If row_count <> col_count Then
        upper_left.Value = OPNAME
        upper_left.Cells(1, 1 + HSPACE).Value = "Not a square matrix."
        Exit Sub
    End If
    
    'Dump the result in the sheet.
    Set dump_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address, upper_left.Cells(row_count, col_count + HSPACE).Address)
    upper_left.Value = OPNAME & CStr(nth)
    dump_range.Value = npow(matrix_range.Value, nth)
End Sub

Private Function npow(mat_A As Variant, n As Integer) As Variant
    'Return the result of taking the n-th power of mat_A.
    Dim y As Variant
    Dim n_rows As Long, n_cols As Long
    
    n_rows = UBound(mat_A, 1) - LBound(mat_A, 1) + 1
    n_cols = UBound(mat_A, 2) - LBound(mat_A, 2) + 1
    
    'Check whether mat_A is a square matrix.
    If n_rows <> n_cols Then
        Debug.Print "Not a square matrix."
        Exit Function
    End If
    
    If n < 0 Then
        Debug.Print "Only positive integer exponents allowed."
        Exit Function
    End If
    If n = 0 Then
        npow = eye(n_rows)
        Exit Function
    End If
    
    y = eye(n_rows)
    Do While n > 1
        If n Mod 2 = 1 Then
            y = dot(mat_A, y)
            n = n - 1
        End If
        mat_A = dot(mat_A, mat_A)
        n = n / 2
    Loop
    npow = dot(mat_A, y)
End Function

Sub test_call_npow()
    Dim mat_A(1 To 3, 1 To 3) As Double, mat_B(1 To 3, 1 To 3) As Double
    Dim i As Integer, j As Integer, k As Integer, n As Integer
    Dim result As Variant
    
    'Matrix to be exponentiated.
    k = 1
    For i = 1 To 3
        For j = 1 To 3
            mat_A(i, j) = k
            k = k + 1
        Next j
    Next i
    
    'Call npow
    n = -1
    result = npow(mat_A, n)
    For i = 1 To 3
        For j = 1 To 3
            Debug.Print result(i, j)
        Next j
    Next i
End Sub

Private Function eye(n As Long) As Variant
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

Sub test_eye()
    Dim n As Long: n = 3
    Dim i As Long, j As Long
    Dim id As Variant
    
    id = eye(n)
    For i = 1 To n
        For j = 1 To n
            Debug.Print id(i, j)
        Next j
    Next i
End Sub
