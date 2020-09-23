VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "MIDAR's 2D Rotation Lesson"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   150
      Top             =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    ' I only placed this here, to force you to read it.
    ' You can also set this property on the form (or picturebox) directly.
    Me.AutoRedraw = True
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.Caption = App.Comments & "  x:" & Int(X) & "  y:" & Int(Y)
    
    Call MouseMoved(X, Y)
    
End Sub


Private Sub Form_Resize()

    ' Square-up the form
    Dim sngAspectRatio As Single
    sngAspectRatio = Me.Width / Me.Height


    ' Change the forms coordinate system
    '       Form's new Top-Left       (-50,-50)
    '       Form's new Bottom-Right   (50,50)
    Me.ScaleTop = -50
    Me.ScaleHeight = 100
    Me.ScaleLeft = -50 * sngAspectRatio
    Me.ScaleWidth = 100 * sngAspectRatio
    
End Sub


Private Sub Timer1_Timer()

    ' Clear Screen
    Me.Cls
    
    ' Do Game Stuff like, AI, Keyboard Handling and Drawing all graphics.
    Call DoMainGameLoop

    ' At the end of this timer event, the screen will automatically get updated - no flicker!
    
End Sub


