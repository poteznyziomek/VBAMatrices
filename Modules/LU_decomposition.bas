Attribute VB_Name = "LU_decomposition"
Option Explicit

Const VSPACE As Integer = 1 'No. of empty rows between used range and result.
Const HSPACE As Integer = 1 'No. of empty cols between op name and result.
Const OPNAME As String = "LU decomposition"

Public Sub LU(Optional matrix_range As Variant, Optional upper_left As Variant)
    'Call the function for LU-decomposing a matrix_range and draw the result in an appropriate sheet.
    Dim row_count As Long, col_count As Long
    Dim LU_arr As Variant
    Dim sheet As Worksheet, L_range As Range, U_range As Range
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
        
    'Dimensions of the matrix_range
    row_count = matrix_range.Rows.Count
    col_count = matrix_range.Columns.Count
    
    'Define ranges for drawing
    Set L_range = sheet.Range(upper_left.Cells(1, 1 + HSPACE).Address, upper_left.Cells(row_count, col_count + HSPACE).Address)
    Set U_range = sheet.Range(upper_left.Cells(1, upper_left.Column + col_count + HSPACE + 1).Address, upper_left.Cells(row_count, 2 * col_count + HSPACE + 1).Address)
    
    'Check if matrix_range is a square matrix
    If row_count <> col_count Then
        upper_left.Value = OPNAME
        upper_left.Cells(1, 1 + HSPACE).Value = "Not a square matrix."
        Exit Sub
    Else
        upper_left.Value = OPNAME
    End If
    
    'Transfer Values
    LU_arr = banachiewicz_lu(matrix_range.Value)
    L_range.Value = LU_arr(0)
    U_range.Value = LU_arr(1)
End Sub

Private Function banachiewicz_lu(mat As Variant) As Variant
    'Carry out Banachiewicz LU decomposition.
    Dim row_count As Long, col_count As Long, i As Long, j As Long, k As Long
    Dim L_array() As Double, U_array() As Double
    
    row_count = UBound(mat, 1)
    col_count = UBound(mat, 2)
    ReDim L_array(1 To row_count, 1 To col_count)
    ReDim U_array(1 To row_count, 1 To col_count)
    
    'Populate L with zeros and U with ones on the main diagonal
    For i = 1 To row_count
        For j = 1 To col_count
            L_array(i, j) = 0
            If i = j Then
                U_array(i, j) = 1
            Else
                U_array(i, j) = 0
            End If
        Next j
    Next i
    
    For i = 1 To col_count 'or row_count
        L_array(i, 1) = mat(i, 1)
    Next i
    For j = 2 To col_count
        
        U_array(1, j) = mat(1, j) / L_array(1, 1)
    Next j
    For j = 2 To col_count
        'a)
        For i = 2 To j - 1
            U_array(i, j) = 0
            For k = 1 To i - 1
                U_array(i, j) = U_array(i, j) + L_array(i, k) * U_array(k, j)
            Next k
            U_array(i, j) = mat(i, j) - U_array(i, j)
            U_array(i, j) = U_array(i, j) / L_array(i, i)
        Next i
        'b)
        For i = j To col_count
            L_array(i, j) = 0
            For k = 1 To j - 1
                L_array(i, j) = L_array(i, j) + L_array(i, k) * U_array(k, j)
            Next k
            L_array(i, j) = mat(i, j) - L_array(i, j)
        Next i
    Next j
    banachiewicz_lu = Array(L_array, U_array)
End Function
