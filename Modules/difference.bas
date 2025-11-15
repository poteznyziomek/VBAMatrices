Attribute VB_Name = "difference"
Option Explicit

Const VSPACE As Integer = 1 'No. of empty rows between used range and result.
Const HSPACE As Integer = 1 'No. of empty cols between op name and result.
Const OPNAME As String = "Difference."

Sub mat_sub(Optional matrix_range As Variant, Optional upper_left As Variant)
    'Call the matrix subtraction function and draw it in an appropriate sheet.
    Dim i As Integer, j As Integer
    Dim dump_range As Range
    Dim row_count As Long, col_count As Long
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
    
    'Check whether there are two matrices.
    If matrix_range.Areas.Count <> 2 Then
        upper_left.Value = OPNAME
        upper_left.Cells(1, 1 + HSPACE).Value = "Invalid number of matrices."
        Exit Sub
    End If
    
    'Check whether subtraction is possible.
    With matrix_range
        If .Areas(1).Rows.Count <> .Areas(2).Rows.Count Then
            'Debug.Print "Not compatible for subtraction."
            upper_left.Value = OPNAME
            upper_left.Cells(1, 1 + HSPACE).Value = "Invalid dimensions."
            Exit Sub
        End If
        If .Areas(1).Columns.Count <> .Areas(2).Columns.Count Then
            'Debug.Print "Not compatible for subtraction."
            upper_left.Value = OPNAME
            upper_left.Cells(1, 1 + HSPACE).Value = "Invalid dimensions."
            Exit Sub
        End If
    End With
    
    'Dump the result in the sheet.
    row_count = matrix_range.Rows.Count
    col_count = matrix_range.Columns.Count
    Set dump_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address, upper_left.Cells(row_count, col_count + HSPACE).Address)
    upper_left.Value = OPNAME
    dump_range.Value = subtract_two(matrix_range.Areas(1).Value, matrix_range.Areas(2).Value)
End Sub

Function subtract_two(mat_A As Variant, mat_B As Variant) As Variant
    'Subtract mat_B from mat_A.
    Dim i As Integer, j As Integer
    Dim m0 As Integer, m As Integer, n0 As Integer, n As Integer
    Dim result_mat() As Variant
    
    m0 = LBound(mat_A, 1): m = UBound(mat_A, 1): n0 = LBound(mat_A, 2): n = UBound(mat_A, 2)
    ReDim result_mat(m0 To m, n0 To n)
    For i = m0 To m
        For j = n0 To n
            result_mat(i, j) = mat_A(i, j) - mat_B(i, j)
        Next j
    Next i
    subtract_two = result_mat
End Function
