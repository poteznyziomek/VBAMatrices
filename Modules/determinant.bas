Attribute VB_Name = "determinant"
Option Explicit

Const VSPACE As Integer = 1 'No. of empty rows between used range and result.
Const HSPACE As Integer = 1 'No. of empty cols between op name and result.
Const OPNAME As String = "Determinant"

Sub mat_det(Optional matrix_range As Variant, Optional upper_left As Variant)
    'Call the determinant function and draw the result in an appropriate sheet.
    Dim result As Variant
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
    
    'Check whether there is a single matrix.
    If matrix_range.Areas.Count > 1 Then
        upper_left.Value = OPNAME
        upper_left.Cells(1, 1 + HSPACE).Value = "Invalid number of matrices."
        Exit Sub
    End If
    
    'Check whether matrix is square.
    If matrix_range.Rows.Count <> matrix_range.Columns.Count Then
        upper_left.Value = OPNAME
        upper_left.Cells(1, 1 + HSPACE).Value = "Invalid dimensions."
        Exit Sub
    End If
    
    'Calculate the determinant.
    result = lu_det(matrix_range.Value)
    
    'Dump the result in the sheet.
    upper_left.Value = OPNAME
    upper_left.Cells(1, 1 + HSPACE).Value = result
End Sub

Function lu_det(mat_A As Variant) As Variant
    'Return the determinant of the matrix if it exists, else string.
    Dim LU_p_sign As Variant
    Dim i As Integer
    
    LU_p_sign = gauss_pp(mat_A) 'Holds three objects as name suggests.
    
    '
    If TypeName(LU_p_sign) = "Integer" Then
        lu_det = "Singular matrix"
        Exit Function
    End If
    
    'Calculate the product of the entries on the main diagonal of U.
    lu_det = 1
    For i = LBound(LU_p_sign(0), 1) To UBound(LU_p_sign(0), 1)
        lu_det = lu_det * LU_p_sign(0)(LU_p_sign(1)(i), i)
    Next i
    'Adjust for the number of swaps.
    lu_det = LU_p_sign(2) * lu_det
End Function
