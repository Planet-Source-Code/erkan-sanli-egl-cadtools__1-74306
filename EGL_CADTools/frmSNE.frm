VERSION 5.00
Begin VB.Form frmSNE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nearest point on the ellipse"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   417
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
Attribute VB_Name = "frmSNE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------
'A Point on the ellipse
'------------------------

Option Explicit

Dim idx         As Byte         'Selected point index
Dim E1          As tELLIPSE
Dim HA          As POINT        'Major axis handel
Dim HB          As POINT        'Minor axis handel
Dim FP          As POINT        'Float point

Private Sub Form_Load()
    
    E1 = SetEllipse(0, 0, 5, 3)
    HA = SetPoint(E1.C.X + E1.A, E1.C.Y)
    HB = SetPoint(E1.C.X, E1.C.Y + E1.B)
    CanvasRescale picCanvas
    Redraw

End Sub


Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        idx = 0
        If GrabPoint(X, Y, HA) Then idx = 1
        If GrabPoint(X, Y, HB) Then idx = 2
    End If
 
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        Select Case idx
            Case 1
                HA.X = Abs(X)
                E1.A = HA.X
            Case 2
                HB.Y = Abs(Y)
                E1.B = HB.Y
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
    DrawPointRect picCanvas, HA, 1
    DrawPointRect picCanvas, HB, 2
    DrawEllipse picCanvas, E1
    NearestPointOnTheEllipse E1, F, FP, NP
    If F Then picCanvas.Circle (NP.X, NP.Y), Mag

End Sub
