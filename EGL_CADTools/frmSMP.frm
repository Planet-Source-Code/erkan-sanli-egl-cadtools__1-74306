VERSION 5.00
Begin VB.Form frmSMP 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Midpoint of the line"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   571
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   120
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   120
      Width           =   6000
   End
   Begin VB.Label lblCurPos 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblCurPos"
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmSMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------
'Midpoint of the line
'--------------------

Option Explicit

Dim idx         As Byte     'Selected point index
Dim L1          As tLINE

Private Sub Form_Load()
    
    L1 = SetLine(5, 3, -7, -4)
    CanvasRescale picCanvas
    Redraw

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        idx = 0
        If GrabPoint(X, Y, L1.P1) Then idx = 1
        If GrabPoint(X, Y, L1.P2) Then idx = 2
    End If
 
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        Select Case idx
            Case 1: L1.P1 = SetPoint(X, Y)
            Case 2: L1.P2 = SetPoint(X, Y)
        End Select
        Redraw
    End If
    lblCurPos.Caption = "X:" & X & "  Y:" & Y
   
End Sub

Private Sub Redraw()
    
    Dim MP  As POINT    'Mid point
    
    CanvasRedraw picCanvas
    DrawPointRect picCanvas, L1.P1, 1
    DrawPointRect picCanvas, L1.P2, 2
    DrawLine picCanvas, L1
    MidPoint L1, MP
    picCanvas.Circle (MP.X, MP.Y), Mag
    
End Sub

