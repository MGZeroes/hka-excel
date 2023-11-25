Attribute VB_Name = "HelperFunctions"
Public Enum MovementDirection
'
' MovementDirection Enum
' Movement Enum for left, right or no movement
'
    NoMovement
    MoveLeft
    MoveRight
End Enum

Public Function Delay(milliseconds As Long)
'
' Delay Function
' Delays execution for a number of milliseconds.
'
    Dim endTime As Double
    endTime = Timer + (milliseconds / 1000)

    Do While Timer < endTime
        DoEvents  ' Yield to other processes and events.
    Loop
End Function

Public Function ValidateSelection(expectedType As String) As Boolean
    '
    ' ValidateSelection Function
    ' Returns True if the current selection is of the expected type,
    ' False otherwise.
    '
    ValidateSelection = (TypeName(Selection) = expectedType)
End Function

Public Function CheckSelection(expectedType As String) As Boolean
'
' CheckSelection Function
' Check if the current selection is of the expected type.
' Send Message if not
'
    If Not ValidateSelection(expectedType) Then
        MsgBox "Please select a " & expectedType & "."
        CheckSelection = False
        Exit Function
    End If
    CheckSelection = True
End Function
