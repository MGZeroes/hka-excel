VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectionForm 
   Caption         =   "Selektieren und Einfügen"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465.001
   OleObjectBlob   =   "SelectionForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "SelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function UpdateTextBoxRangeValue()
'
' UpdateTextBoxRangeValue Function
' Update the Textbox to contain the current Selection
'
    If Not CheckSelection("Range") Then Exit Function

    On Error Resume Next ' Enable error handling
    TextBoxRange.Value = Selection.Address
    On Error GoTo 0 ' Turn off error handling
End Function

Private Function SelectTextBoxRangeValue()
'
' SelectTextBoxRangeValue Function
' Select the Range in the given TextBox
'
    On Error Resume Next ' Enable error handling
    Range(TextBoxRange.Value).Select
    On Error GoTo 0 ' Turn off error handling
End Function

Private Function InitComboBoxNames()
'
' InitComboBoxNames Function
' Fill the given ComboBox with all the defined names
'
    Dim nm As name

    ' Clear existing items
    ComboBoxNames.Clear

    ' Loop through each name in the workbook
    For Each nm In ThisWorkbook.Names
        ComboBoxNames.AddItem nm.name
    Next nm
End Function

''' INIT AND EXIT
Private Sub UserForm_Initialize()
    UpdateTextBoxRangeValue
    InitComboBoxNames
End Sub

Private Sub ButtonClose_Click()
    Unload Me
End Sub

''' INSERT AREA
Private Sub ButtonPaste_Click()
'
' ButtonPaste_Click Event
' Paste the given TextBox value in every selected cell
'
    If Not CheckSelection("Range") Then Exit Sub
    Selection.Value = TextBoxValue.Value
End Sub

Private Sub ButtonDelete_Click()
'
' ButtonDelete_Click Event
' Empty every selected cell
'
    If Not CheckSelection("Range") Then Exit Sub
    Selection.Value = ""
End Sub

Private Sub ButtonMultiply_Click()
'
' ButtonMultiply_Click Event
' Multiply every selected cell with a numeric value
' by the value in the given TextBox
'
    If Not CheckSelection("Range") Then Exit Sub

    ' Iterate through every selected cell
    Dim cell As Range
    For Each cell In Selection
        ' Multiply if the cell contains a numeric value
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            ' Multiply by the TextBox value
            cell.Value = cell.Value * TextBoxMultiply.Value
        End If
    Next cell
End Sub

Private Sub TextBoxMultiply_Change()
'
' TextBoxMultiply_Change Event
' Validate if the value of the TextBox is numeric
'
    ' Set an empty value to 0
    If TextBoxMultiply.Value = "" Then
        TextBoxMultiply.Value = "0"
    ' Check if the TextBox value is numeric
    ElseIf Not IsNumeric(TextBoxMultiply.Value) Then
        MsgBox "Only numeric values are allowed", vbExclamation
        TextBoxMultiply.Value = "0"
    End If
End Sub

''' SELECTION AREA
Private Sub OptionRange_Change()
'
' OptionRange_Change Event
' Enables the TextBox for setting the Selection Range
'
    TextBoxRange.Enabled = OptionRange.Value
    If TextBoxRange.Enabled = True Then
        UpdateTextBoxRangeValue
        SelectTextBoxRangeValue
    End If
End Sub

Private Sub TextBoxRange_Change()
'
' TextBoxRange_Change Event
' Set the TextBox value to the Selection Range
'
    If TextBoxRange.Enabled = True Then
        SelectTextBoxRangeValue
    End If
End Sub

Private Sub OptionPositiveRelative_Change()
'
' OptionPositiveRelative_Change Event
' Set the Selection Range to a 4x4 grid from the Selection Cell
'
    If Not CheckSelection("Range") Then Exit Sub
    
    If OptionPositiveRelative.Value = True Then
        Selection.Range(Cells(1, 1), Cells(4, 4)).Select
    End If
End Sub

Private Sub OptionNegativeRelative_Change()
'
' OptionNegativeRelative_Change Event
' Set the Selection Range to a -4x-4 grid from the Selection Cell
'
    If Not CheckSelection("Range") Then Exit Sub
    
    If OptionNegativeRelative.Value = True Then
        On Error Resume Next  ' Enable error handling
        Selection.Offset(-3, -3).Range(Cells(1, 1), Cells(4, 4)).Select
        On Error GoTo 0  ' Turn off error handling
    End If
End Sub

Private Sub OptionName_Change()
'
' OptionName_Change Event
' Enable the ComboBox for picking the names
'
    ComboBoxNames.Enabled = OptionName.Value
End Sub

Private Sub ComboBoxNames_Change()
'
' ComboBoxNames_Change Event
' Set a defined name as Selection Range
'
    If ComboBoxNames.Enabled = True Then
        Range(ComboBoxNames.Value).Select
    End If
End Sub
