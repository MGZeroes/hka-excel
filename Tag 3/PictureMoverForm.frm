VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PictureMoverForm 
   Caption         =   "Bild Bewegen"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "PictureMoverForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "PictureMoverForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MoveDirection As MovementDirection

Private Function MoveImage()
    ' Move until the left/right border has been reached
    Do While ((MoveDirection = MoveLeft And ImageMovable.Left > 12) Or _
              (MoveDirection = MoveRight And ImageMovable.Left < (Me.Width - ImageMovable.Width - 24)))
        
        ' Abort when closing the Form
        If Me.Controls.Count = 0 Or MoveDirection = NoMovement Then Exit Function

        ' Move into direction
        If MoveDirection = MoveLeft Then
            ImageMovable.Left = ImageMovable.Left - 5
        ElseIf MoveDirection = MoveRight Then
            ImageMovable.Left = ImageMovable.Left + 5
        End If
        
        ' Delay for smoother movement
        Delay 8
    Loop

    MoveDirection = NoMovement
End Function

Private Sub SpinButtonDirection_SpinUp()
    MoveDirection = MoveRight
    MoveImage
End Sub

Private Sub SpinButtonDirection_SpinDown()
    MoveDirection = MoveLeft
    MoveImage
End Sub

Private Sub ButtonHide_Click()
    ImageMovable.Visible = Not ButtonHide.Value
End Sub

Private Sub ButtonClose_Click()
    MoveDirection = NoMovement
    Unload Me
End Sub
