Attribute VB_Name = "addition"
Option Explicit

Const VSPACE As Integer = 1 'No. of empty rows between used range and result.
Const HSPACE As Integer = 1 'No. of empty cols between op name and result.
Const OPNAME As String = "Sum"

Sub mat_add(Optional matrix_range As Variant, Optional upper_left As Variant)
    'Call the matrix addition function and draw the result in an appropriate sheet.
    Dim i As Integer, j As Integer, k As Integer
    Dim mat_dims() As Integer '(Nrows,Ncols) in mat_dims
    Dim family() As Variant 'holds the matrices
    Dim dump_range As Range
    Dim sheet As Worksheet
    Dim result_mat As Variant
    Dim row_count As Long, col_count As Long
    Set sheet = ActiveSheet
    
    If IsMissing(matrix_range) Then
        If TypeName(Selection) = "Range" Then
            Set matrix_range = Selection
        Else
            MsgBox TypeName(Selection) & " is not Range"
            Exit Sub
        End If
    Else
        'The Range matrix_range is passed by OK button click event.
    End If
    If IsMissing(upper_left) Then
        Set upper_left = sheet.Cells(sheet.UsedRange.Rows.Count + 1 + VSPACE, 1)
    End If
    
    ReDim mat_dims(1 To matrix_range.Areas.Count, 2)
    With matrix_range
        For i = 1 To .Areas.Count
            mat_dims(i, 1) = .Areas(i).Rows.Count
            mat_dims(i, 2) = .Areas(i).Columns.Count
        Next i
    End With
    
    'Check whether addition is possible.
    For i = 1 To matrix_range.Areas.Count - 1
        If mat_dims(i, 1) <> mat_dims(i + 1, 1) _
          Or mat_dims(i, 2) <> mat_dims(i + 1, 2) Then
            upper_left.Value = OPNAME
            upper_left.Cells(1, 1 + HSPACE).Value = "Invalid dimensions."
            Exit Sub
        End If
    Next i
    
    'Populate the family array
    ReDim family(1 To matrix_range.Areas.Count)
    For i = 1 To matrix_range.Areas.Count
        family(i) = matrix_range.Areas(i).Value
    Next i
    
    'Dump the result in the sheet.
    row_count = matrix_range.Rows.Count
    col_count = matrix_range.Columns.Count
    Set dump_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address, upper_left.Cells(row_count, col_count + HSPACE).Address)
    upper_left.Value = OPNAME
    dump_range.Value = add_many(family)
End Sub

Private Function add_many(family As Variant) As Variant
    'Sum the matrices in the family.
    Dim i As Integer, j As Integer, k As Integer
    Dim n_row As Integer, n_col As Integer, family_count As Integer
    n_row = UBound(family(1), 1): n_col = UBound(family(1), 2)
    Dim result_mat() As Variant
    ReDim result_mat(1 To n_row, 1 To n_col)
    
    For i = 1 To n_row
        For j = 1 To n_col
            result_mat(i, j) = 0
            For k = LBound(family) To UBound(family)
                result_mat(i, j) = result_mat(i, j) + family(k)(i, j)
            Next k
        Next j
    Next i
    add_many = result_mat
End Function
