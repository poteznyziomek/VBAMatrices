Attribute VB_Name = "multiplication"
Option Explicit

Const VSPACE As Integer = 1 'No. of empty rows between used range and result.
Const HSPACE As Integer = 1 'No. of empty cols between op name and result.
Const OPNAME As String = "Product"

Sub mat_mul(Optional matrix_range As Variant, Optional upper_left As Variant)
    'Call a function to calculate a product of multiple matrices and draw the result in an appropriate sheet.
    Dim i As Integer ', j As Integer, k As Integer
    Dim mat_dims() As Integer '(Nrows,Ncols) in mat_dims
    Dim family() As Variant
    Dim dump_range As Range
    Dim sheet As Worksheet
    Dim row_count As Long, col_count As Long, width As Long, height As Long
    Dim arr_1x1(1 To 1, 1 To 1)
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
    
    ReDim mat_dims(1 To matrix_range.Areas.Count, 2)
    
    With matrix_range
        For i = 1 To .Areas.Count
            mat_dims(i, 1) = .Areas(i).Rows.Count
            mat_dims(i, 2) = .Areas(i).Columns.Count
        Next i
    End With
    
    'Check whether multiplication is possible.
    For i = 1 To matrix_range.Areas.Count - 1
        If mat_dims(i, 2) <> mat_dims(i + 1, 1) Then
            upper_left.Value = OPNAME
            upper_left.Cells(1, 1 + HSPACE).Value = "Invalid dimensions."
            Exit Sub
        End If
    Next i
    
    'Pack the matrices into a family.
    ReDim family(1 To matrix_range.Areas.Count)
    For i = 1 To matrix_range.Areas.Count
        If mat_dims(i, 1) = 1 And mat_dims(i, 2) = 1 Then 'A single cell needs to be packed as a 1 by 1 array.
            arr_1x1(1, 1) = matrix_range.Areas(i).Value
            family(i) = arr_1x1
        Else
            family(i) = matrix_range.Areas(i).Value
        End If
    Next i
    
    height = mat_dims(1, 1)
    width = mat_dims(UBound(mat_dims, 1), 2)
    Set dump_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address, upper_left.Cells(height, width + HSPACE).Address)
    upper_left.Value = OPNAME
    dump_range.Value = dot_many(family)
End Sub

Function dot(mat_A As Variant, mat_B As Variant) As Variant
    'Return the result of multiplying two matrices.
    Dim mA0 As Integer, mA As Integer, nA0 As Integer, nA As Integer 'dimensions of mat_A
    Dim mB0 As Integer, mB As Integer, nB0 As Integer, nB As Integer 'dimensions of mat_B
    Dim i As Integer, j As Integer, k As Integer
    Dim result_mat() As Double
    
    mA0 = LBound(mat_A, 1): mA = UBound(mat_A, 1): nA0 = LBound(mat_A, 2): nA = UBound(mat_A, 2)
    mB0 = LBound(mat_B, 1): mB = UBound(mat_B, 1): nB0 = LBound(mat_B, 2): nB = UBound(mat_B, 2)
    
    'Check for compatibility.
    If nA - nA0 + 1 <> mB - mB0 + 1 Or nA0 <> mB0 Or nA <> mB Then
        Debug.Print "Incompatible matrices for mul."
        Exit Function
    End If
    
    ReDim result_mat(mA0 To mA, nB0 To nB)
    
    'Carry out the multiplication of mat_A and mat_B.
    For i = mA0 To mA
        For j = nB0 To nB
        result_mat(i, j) = 0
            For k = mB0 To mB
                result_mat(i, j) = result_mat(i, j) + mat_A(i, k) * mat_B(k, j)
            Next k
        Next j
    Next i
    dot = result_mat
End Function

Private Function dot_many(family As Variant) As Variant
    'Multiply matrices from family.
    Dim mat_dims() As Integer, temp_mat() As Double, result_mat As Variant
    Dim i As Integer, j As Integer
    
    ReDim mat_dims(LBound(family) To UBound(family), 1 To 4)
    For i = LBound(family) To UBound(family)
        mat_dims(i, 1) = LBound(family(i), 1) 'm0^(i)
        mat_dims(i, 2) = UBound(family(i), 1) 'm^(i)
        mat_dims(i, 3) = LBound(family(i), 2) 'n0^(i)
        mat_dims(i, 4) = UBound(family(i), 2) 'n^(i)
    Next i
    
    'Carry out the multiplication of multiple matrices
    ReDim result_mat(mat_dims(1, 1) To mat_dims(1, 2), mat_dims(1, 3) To mat_dims(1, 4))
    result_mat = family(1)
    For i = LBound(family) To UBound(family) - 1
        'Multiply ith and (i+1)th matrices. The ith is stored in result_mat.
        ReDim temp_mat(LBound(result_mat, 1) To UBound(result_mat, 1), mat_dims(i + 1, 3) To mat_dims(i + 1, 4))
        temp_mat = dot(result_mat, family(i + 1))
        ReDim result_mat(LBound(temp_mat, 1) To UBound(temp_mat, 1), LBound(temp_mat, 2) To UBound(temp_mat, 2))
        result_mat = temp_mat
    Next i
    dot_many = result_mat
End Function

Function dotOBSOLETE(mat_A As Variant, mat_B As Variant) As Variant
    'OBSOLETE. Use dot instead.
    'Return the result of multiplying two matrices.
    Dim row_A As Integer, col_A As Integer, row_B As Integer, col_B As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim result_mat() As Double
    row_A = UBound(mat_A, 1): col_A = UBound(mat_A, 2)
    row_B = UBound(mat_B, 1): col_B = UBound(mat_B, 2)

    'Check for compatibility.
    If col_A <> row_B Then
        Exit Function
    End If

    ReDim result_mat(1 To row_A, 1 To col_B)

    'Carry out the multiplication of mat_A and mat_B.
    For i = 1 To row_A
        For j = 1 To col_B
        result_mat(i, j) = 0
            For k = 1 To col_A
                result_mat(i, j) = result_mat(i, j) + mat_A(i, k) * mat_B(k, j)
            Next k
        Next j
    Next i
    dotOBSOLETE = result_mat
End Function

Private Function dot_manyOBSOLETE(family As Variant) As Variant
    'OBSOLETE. Use dot_many instead.
    'Multiply matrices from family.
    Dim mat_dims() As Integer, temp_mat() As Double, result_mat As Variant
    Dim i As Integer, j As Integer
    
    ReDim mat_dims(LBound(family) To UBound(family), 2)
    For i = LBound(family) To UBound(family)
        mat_dims(i, 1) = UBound(family(i), 1)
        mat_dims(i, 2) = UBound(family(i), 2)
    Next i
    
    'Carry out the multiplication of multiple matrices
    ReDim result_mat(1 To mat_dims(1, 1), 1 To mat_dims(1, 2))
    result_mat = family(1)
    For i = LBound(family) To UBound(family) - 1
        'Multiply ith and (i+1)th matrices. The ith is stored in result_mat.
        ReDim temp_mat(1 To mat_dims(i, 1), 1 To mat_dims(i + 1, 2))
        temp_mat = dot(result_mat, family(i + 1))
        ReDim result_mat(1 To mat_dims(i, 1), 1 To mat_dims(i + 1, 2))
        result_mat = temp_mat
    Next i
    dot_manyOBSOLETE = result_mat
End Function

