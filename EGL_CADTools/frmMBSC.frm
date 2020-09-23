VERSION 5.00
Begin VB.Form frmMBSC 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bisector 2 lines"
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
Attribute VB_Name = "frmMBSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------
'Draw a circle from 3 lines
'--------------------

Option Explicit

Dim idx         As Byte     'Selected point index
Dim L1          As tLINE
Dim L2          As tLINE
Dim L3          As tLINE

Private Sub Form_Load()
    
    L1 = SetLine(5, 3, -7, -4)
    L2 = SetLine(8, -8, -7, -7)
    L3 = SetLine(7, -5, 7, 6)
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
        If GrabPoint(X, Y, L3.P1) Then idx = 5
        If GrabPoint(X, Y, L3.P2) Then idx = 6
    End If
 
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        Select Case idx
            Case 1: L1.P1 = SetPoint(X, Y)
            Case 2: L1.P2 = SetPoint(X, Y)
            Case 3: L2.P1 = SetPoint(X, Y)
            Case 4: L2.P2 = SetPoint(X, Y)
            Case 5: L3.P1 = SetPoint(X, Y)
            Case 6: L3.P2 = SetPoint(X, Y)
        End Select
        Redraw
    End If
    lblCurPos.Caption = "X:" & X & "  Y:" & Y
   
End Sub

Private Sub Redraw()
    
    Dim DC As tCIRCLE  'Dot Circle
        
    picCanvas.DrawStyle = vbSolid
    CanvasRedraw picCanvas
    DrawPointRect picCanvas, L1.P1, 1
    DrawPointRect picCanvas, L1.P2, 2
    DrawPointRect picCanvas, L2.P1, 3
    DrawPointRect picCanvas, L2.P2, 4
    DrawPointRect picCanvas, L3.P1, 5
    DrawPointRect picCanvas, L3.P2, 6
    DrawLine picCanvas, L1
    DrawLine picCanvas, L2
    DrawLine picCanvas, L3
    CircleLLL L1, L2, L3, DC
    picCanvas.DrawStyle = vbDot
    DrawCircle picCanvas, DC
   
End Sub
