VERSION 5.00
Begin VB.Form frmC3P 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Draw a circle from 3 points"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   572
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
Attribute VB_Name = "frmC3P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------
'Equation of a Circle from 3 Points
'------------------------

Option Explicit

Dim idx         As Byte     'Selected point index
Dim P(1 To 3)   As POINT    'Points
Dim C1          As tCIRCLE

Private Sub Form_Load()
    
    P(1) = SetPoint(-5, -5)
    P(2) = SetPoint(-3, 5)
    P(3) = SetPoint(3, 1)
    CanvasRescale picCanvas
    Redraw

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        idx = 0
        If GrabPoint(X, Y, P(1)) Then idx = 1
        If GrabPoint(X, Y, P(2)) Then idx = 2
        If GrabPoint(X, Y, P(3)) Then idx = 3
    End If
 
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        If idx <> 0 Then
            P(idx) = SetPoint(X, Y)
            Redraw
        End If
    End If
    lblCurPos.Caption = "X:" & X & "  Y:" & Y

End Sub

Private Sub Redraw()
    
    picCanvas.DrawStyle = vbSolid
    CanvasRedraw picCanvas
    DrawPointRect picCanvas, P(1), 1
    DrawPointRect picCanvas, P(2), 2
    DrawPointRect picCanvas, P(3), 3
    Circle3Points P, C1
    picCanvas.DrawStyle = vbDot
    DrawCircle picCanvas, C1
   
End Sub
