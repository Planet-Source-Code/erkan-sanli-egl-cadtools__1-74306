VERSION 5.00
Begin VB.Form frmCLCR 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Draw a circle from a line and a circle tangents and radius"
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
   Begin VB.TextBox txtr 
      Height          =   285
      Left            =   6720
      TabIndex        =   2
      Text            =   "2"
      Top             =   5640
      Width           =   975
   End
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
   Begin VB.Label Label1 
      Caption         =   "Radius"
      Height          =   255
      Left            =   6360
      TabIndex        =   4
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      Top             =   5640
      Width           =   375
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
Attribute VB_Name = "frmCLCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------
'Draw a circle a line and a circle tangents and radius
'--------------------

Option Explicit

Dim idx         As Byte     'Selected point index
Dim L1          As tLINE
Dim C1          As tCIRCLE
Dim HR          As POINT

Private Sub Form_Load()

    L1 = SetLine(9, 9, -7, 1)
    C1 = SetCircle(4, 1, 3)
    HR = SetPoint(C1.C.X + C1.R, C1.C.Y)
    CanvasRescale picCanvas
    Redraw

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        idx = 0
        If GrabPoint(X, Y, L1.P1) Then idx = 1
        If GrabPoint(X, Y, L1.P2) Then idx = 2
        If GrabPoint(X, Y, C1.C) Then idx = 3
        If GrabPoint(X, Y, HR) Then idx = 4
    End If
 
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        Select Case idx
            Case 1
                L1.P1 = SetPoint(X, Y)
            Case 2
                L1.P2 = SetPoint(X, Y)
            Case 3
                C1.C = SetPoint(X, Y)
                HR = SetPoint(X + C1.R, Y)
            Case 4
                HR = SetPoint(X, Y)
                C1.R = Distance(C1.C.X, C1.C.Y, HR.X, HR.Y)
        End Select
        Redraw
    End If
    lblCurPos.Caption = "X:" & X & "  Y:" & Y
   
End Sub

Private Sub txtr_Change()
    
    Redraw

End Sub

Private Sub Redraw()
    
    Dim F(1 To 2)   As Boolean  'Flags
    Dim DC(1 To 2)  As tCIRCLE  'Dot Circle
    Dim R           As Single
    
    picCanvas.DrawStyle = vbSolid
    CanvasRedraw picCanvas
    DrawPointRect picCanvas, L1.P1, 1
    DrawPointRect picCanvas, L1.P2, 2
    DrawPointRect picCanvas, C1.C, 3
    DrawPointRect picCanvas, HR, 4
    DrawLine picCanvas, L1
    DrawCircle picCanvas, C1
    R = GetVal(txtr)
    CircleLCR L1, C1, R, F, DC
    picCanvas.DrawStyle = vbDot
    If F(1) Then DrawCircle picCanvas, DC(1)
    If F(2) Then DrawCircle picCanvas, DC(2)

End Sub
