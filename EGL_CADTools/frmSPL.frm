VERSION 5.00
Begin VB.Form frmSPL 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Perpendicular point on the Line"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
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
   Begin VB.Label lblDistance 
      Caption         =   "D:"
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Minimum Distance between a point and a line"
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   720
      Width           =   2175
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
Attribute VB_Name = "frmSPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------
'Perpendicular point on the Line
'-------------------

Option Explicit

Dim idx         As Byte     'Selected point index
Dim L1          As tLINE
Dim FP          As POINT    'Float point

Private Sub Form_Load()
    
    L1 = SetLine(5, 3, -7, -4)
    FP = SetPoint(5, -4)
    CanvasRescale picCanvas
    Redraw

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        idx = 0
        If GrabPoint(X, Y, L1.P1) Then idx = 1
        If GrabPoint(X, Y, L1.P2) Then idx = 2
        If GrabPoint(X, Y, FP) Then idx = 3
    End If
 
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        Select Case idx
            Case 1: L1.P1 = SetPoint(X, Y)
            Case 2: L1.P2 = SetPoint(X, Y)
            Case 3: FP = SetPoint(X, Y)
        End Select
    End If
    Redraw
    lblCurPos.Caption = "X:" & X & "  Y:" & Y

End Sub

Private Sub Redraw()
    
    Dim PP As POINT
    
    picCanvas.DrawStyle = vbSolid
    CanvasRedraw picCanvas
    DrawPointRect picCanvas, L1.P1, 1
    DrawPointRect picCanvas, L1.P2, 2
    DrawPointRect picCanvas, FP, 3
    DrawLine picCanvas, L1
    PerpendicularPointOnTheLine L1, FP, PP
    picCanvas.Circle (PP.X, PP.Y), Mag
    picCanvas.DrawStyle = vbDot
    picCanvas.Line (PP.X, PP.Y)-(FP.X, FP.Y)
    lblDistance.Caption = "D: " & Distance(PP.X, PP.Y, FP.X, FP.Y)
    
End Sub
