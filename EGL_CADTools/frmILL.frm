VERSION 5.00
Begin VB.Form frmILL 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Intersection point of two lines"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   419
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
Attribute VB_Name = "frmILL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------
'Intersection point of two lines
'----------------------

Option Explicit

Dim idx         As Byte     'Selected point index
Dim L1          As tLINE
Dim L2          As tLINE

Private Sub Form_Load()
    
    L1 = SetLine(-5, -5, 5, 5)
    L2 = SetLine(-3, 5, 5, -3)
    CanvasRescale picCanvas
    Redraw

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        idx = 0
        If GrabPoint(X, Y, L1.P1) Then idx = 1
        If GrabPoint(X, Y, L1.P2) Then idx = 2
        If GrabPoint(X, Y, L2.P1) Then idx = 3
        If GrabPoint(X, Y, L2.P2) Then idx = 4
    End If
 
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        Select Case idx
            Case 1: L1.P1 = SetPoint(X, Y)
            Case 2: L1.P2 = SetPoint(X, Y)
            Case 3: L2.P1 = SetPoint(X, Y)
            Case 4: L2.P2 = SetPoint(X, Y)
        End Select
        Redraw
    End If
    lblCurPos.Caption = "X:" & X & "  Y:" & Y

End Sub

Private Sub Redraw()
    
    Dim IP  As POINT    'Intersection point
    
    CanvasRedraw picCanvas
    DrawPointRect picCanvas, L1.P1, 1
    DrawPointRect picCanvas, L1.P2, 2
    DrawPointRect picCanvas, L2.P1, 3
    DrawPointRect picCanvas, L2.P2, 4
    DrawLine picCanvas, L1
    DrawLine picCanvas, L2
    IntersectionLineLine L1, L2, IP
    picCanvas.Circle (IP.X, IP.Y), Mag

End Sub
