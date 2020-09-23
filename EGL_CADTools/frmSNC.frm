VERSION 5.00
Begin VB.Form frmSNC 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nearest point on the Circle"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   418
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   568
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
Attribute VB_Name = "frmSNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------
'A Point on the Circle
'------------------------

Option Explicit

Dim idx         As Byte     'Selected point index
Dim C1          As tCIRCLE
Dim HR          As POINT    'Handel radius
Dim FP          As POINT    'Float point

Private Sub Form_Load()
    
    C1 = SetCircle(3, 1, 3)
    HR = SetPoint(C1.C.X + C1.R, C1.C.Y)
    CanvasRescale picCanvas
    Redraw

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        idx = 0
        If GrabPoint(X, Y, C1.C) Then idx = 1
        If GrabPoint(X, Y, HR) Then idx = 2
    End If
 
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        Select Case idx
            Case 1
                C1.C = SetPoint(X, Y)
                HR = SetPoint(X + C1.R, Y)
            Case 2
                HR = SetPoint(X, Y)
                C1.R = Distance(C1.C.X, C1.C.Y, HR.X, HR.Y)
        End Select
    Else
        FP = SetPoint(X, Y)
    End If
    Redraw
    lblCurPos.Caption = "X:" & X & "  Y:" & Y

End Sub

Private Sub Redraw()

    Dim F       As Boolean  'On the circle Flag
    Dim NP      As POINT    'Nearest point

    CanvasRedraw picCanvas
    DrawPointRect picCanvas, C1.C, 1
    DrawPointRect picCanvas, HR, 2
    DrawCircle picCanvas, C1
    NearestPointOnTheCircle C1, F, FP, NP
    If F Then picCanvas.Circle (NP.X, NP.Y), Mag
    
End Sub
