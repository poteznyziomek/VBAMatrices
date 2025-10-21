VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MatOps 
   Caption         =   "Matrix ops"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9435.001
   OleObjectBlob   =   "MatOps.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MatOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const REPORTTITLE As String = "Report "
Const VSPACE As Integer = 1 'No. of empty rows between used range and result.
Const HSPACE As Integer = 1 'No. of empty cols between op name and result.

Public user_range As Range

Private Sub cbCancel_Click()
    Unload Me
End Sub

Private Sub cbOK_Click()
    'Generates a report.
    'Creates a new sheet called Report N and makes it the active sheet.
    'For each checked box calls an appropriate procedure (sub).
    
    'IMPORTANT: Check whether value in textbox for NPow is valid.
    
    Dim report_sheet As Worksheet, report_sheet_name As String
    Dim upper_left As Range
    Dim op_mode As String 'Single or Multiple matrices
    Dim op_control As Variant 'checkbox for operation
    Dim control As Variant
    Dim nth As Integer
    
    'Create the report sheet
    Set report_sheet = ActiveWorkbook.Sheets.Add( _
        After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    report_sheet.name = create_report_sheet_name(REPORTTITLE)
    
    'Define upper left corner for drawing in the report sheet
    Set upper_left = report_sheet.Cells(1, 1)

    'Mode: single or multiple
    For Each control In frNoOfMatrices.Controls
        If control.Value = True Then
            op_mode = control.name
        End If
    Next control
    
    Select Case op_mode
    Case obSingleMat.name
        'Loop through all the checkboxes in the single matrix frame.
        For Each op_control In frSingle.Controls
            If TypeName(op_control) = "CheckBox" Then
                If op_control.Value = True Then
                    'Call appropriate procedure.
                    If op_control.name = "cbSinglePow" Then
                        nth = sbNPowVal.Value
                        Debug.Print nth; TypeName(nth)
                        Application.Run op_control.Tag, nth, user_range, upper_left
                    Else
                        Application.Run op_control.Tag, user_range, upper_left
                    End If
                    Set upper_left = report_sheet.Cells(report_sheet.UsedRange.Rows.Count + 1 + VSPACE, 1)
                End If
            End If
        Next op_control
    Case obMultipleMat.name
        'Loop through all the checkboxes in the single matrix frame.
        For Each op_control In frMultiple.Controls
            If TypeName(op_control) = "CheckBox" Then
                If op_control.Value = True Then
                    'Call appropriate procedure.
                    Application.Run op_control.Tag, user_range, upper_left
                    Set upper_left = report_sheet.Cells(report_sheet.UsedRange.Rows.Count + 1 + VSPACE, 1)
                End If
            End If
        Next op_control
    End Select
    
    Unload Me
End Sub

Private Sub cbSelectCells_Click()
    Dim i As Integer
    Dim caption_text As String
    Dim Prompt As String
    Dim Title As Variant
    Dim Default As Variant
    Dim ibType As Variant
    Dim addresses() As String
    Prompt = "Select cells"
    Title = "Specify input"
    Default = user_range.Address
    ibType = 8
    
    'Hide the form
    Me.Hide
    
    On Error GoTo canceled
    Set user_range = Application.InputBox( _
        Prompt:=Prompt, Title:=Title, _
        Default:=Default, Type:=ibType)
    
    'Resume normal error triggering
    On Error GoTo 0
    user_range.Select
    'User selected incompatible number of matrices
    'and sheet areas.
    If obSingleMat.Value And user_range.Areas.Count > 1 Then
        lblSelectedCells.ForeColor = RGB(255, 0, 0) 'Red
        cbOK.Enabled = False
    ElseIf obSingleMat.Value And user_range.Areas.Count = 1 Then
        lblSelectedCells.ForeColor = RGB(0, 0, 0)
        cbOK.Enabled = True
    ElseIf obMultipleMat.Value And user_range.Areas.Count = 1 Then
        lblSelectedCells.ForeColor = RGB(255, 0, 0)
        cbOK.Enabled = False
    Else
        lblSelectedCells.ForeColor = RGB(0, 0, 0) 'Red
        cbOK.Enabled = True
    End If
    
    'Update the lblSelectedCells
    addresses = Split(user_range.Address, ",")
    caption_text = "Selected cells:"
    For i = 0 To mini(arr_len(addresses) - 1, 2)
        caption_text = caption_text & vbCrLf & addresses(i) & ","
    Next i
    If arr_len(addresses) > 3 Then
        caption_text = caption_text & vbCrLf & "..."
    End If
    lblSelectedCells.Caption = caption_text
    
canceled:
    Me.show
End Sub


Private Sub obSingleMat_Click()
    Dim control As Variant
    
    'Disabled controls on frMultiple
    If frMultiple.Enabled Then
        For Each control In frMultiple.Controls
            control.Enabled = False
        Next control
        frMultiple.Enabled = False
    End If
    
    'Enable controls on frSingle
    If Not frSingle.Enabled Then
        For Each control In frSingle.Controls
            control.Enabled = True
        Next control
        frSingle.Enabled = True
    End If
    
    'Disable/enable OK button
    If user_range.Areas.Count > 1 Then
        cbOK.Enabled = False
    Else
        cbOK.Enabled = True
    End If
    
    'Color code the Selected Cells label
    If user_range.Areas.Count > 1 Then
        lblSelectedCells.ForeColor = RGB(255, 0, 0)
    Else
        lblSelectedCells.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub obMultipleMat_Click()
    Dim control As Variant
    
    'Disable controls on frSingle
    If frSingle.Enabled Then
        For Each control In frSingle.Controls
            control.Enabled = False
        Next control
        frSingle.Enabled = False
    End If
    
    'Enable controls on frMultiple
    If Not frMultiple.Enabled Then
        For Each control In frMultiple.Controls
            control.Enabled = True
        Next control
        frMultiple.Enabled = True
    End If
    
    'Disable/enable OK button
    If user_range.Areas.Count = 1 Then
        cbOK.Enabled = False
    Else
        cbOK.Enabled = True
    End If
    
    'Color code the Selected Cells label
    If user_range.Areas.Count > 1 Then
        lblSelectedCells.ForeColor = RGB(0, 0, 0)
    Else
        lblSelectedCells.ForeColor = RGB(255, 0, 0)
    End If
End Sub

Private Sub sbNPowVal_Change()
    tbNPowVal.Text = sbNPowVal.Value
End Sub

Private Sub tbNPowVal_Change()
    Dim new_value As Integer
    If IsNumeric(tbNPowVal) Then
        new_value = Val(tbNPowVal.Text)
        If new_value >= sbNPowVal.min And new_value <= sbNPowVal.max Then
            sbNPowVal.Value = new_value
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    'Set properties of the controls.
    'Subject to change.
    
    'The Tag property hold a sub name to be called.
    'Single matrix frame
    cbSingleRank.Tag = "placeholder_sub"
    cbSingleLU.Tag = "LU"
    cbSingleDet.Tag = "placeholder_sub"
    cbSingleNil.Tag = "placeholder_sub"
    cbSingleEigen.Tag = "placeholder_sub"
    cbSinglePow.Tag = "mat_npow"
    'Trigonometric frame
    cbSingleSin.Tag = "placeholder_sub"
    cbSingleCos.Tag = "placeholder_sub"
    cbSingleTan.Tag = "placeholder_sub"
    cbSingleCot.Tag = "placeholder_sub"
    'Hyperbolic frame
    cbSingleSinh.Tag = "placeholder_sub"
    cbSingleCosh.Tag = "placeholder_sub"
    cbSingleTanh.Tag = "placeholder_sub"
    cbSingleCoth.Tag = "placeholder_sub"
    'Inverse trigonometric
    cbSingleArcSin.Tag = "placeholder_sub"
    cbSingleArcCos.Tag = "placeholder_sub"
    cbSingleArcTan.Tag = "placeholder_sub"
    cbSingleArcCot.Tag = "placeholder_sub"
    'Misc. frame
    cbSingleLog.Tag = "placeholder_sub"
    'Multiple matrices frame
    cbMultipleSum.Tag = "mat_add"
    cbMultipleProd.Tag = "mat_mul"
    cbMultipleDiff.Tag = "mat_sub"
    cbMultipleSysLinSolve.Tag = "linear_system_solve"
    
    Set user_range = Selection.Cells(1, 1).CurrentRegion
    user_range.Select
    lblSelectedCells.Caption = "Selected cells:" & vbCrLf & user_range.Address
    
    'Properties of the NPow spin button and text box
    sbNPowVal.min = 0
    sbNPowVal.max = 100
    sbNPowVal.Value = 2
End Sub
