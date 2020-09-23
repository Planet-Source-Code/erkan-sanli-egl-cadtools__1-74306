VERSION 5.00
Begin VB.Form frmMOFL 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Offset Lines (Parallel Lines)"
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
   Begin VB.CheckBox chkShow 
      Caption         =   "1. Offset Line"
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "2. Offset Line"
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   5
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txtod 
      Height          =   285
      Left            =   6720
      TabIndex        =   2
      Text            =   "3"
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
      Caption         =   "Offset distance"
      Height          =   255
      Left            =   6360
      TabIndex        =   4
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      Caption         =   "OD"
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
Attribute VB_Name = "frmMOFL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------
'Offset line
'--------------------

Option Explicit

Dim idx         As Byte     'Selected point index
Dim L1          As tLINE

Private Sub Form_Load()
    
    chkShow(0).Value = vbChecked
    chkShow(1).Value = vbChecked
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
            Case 1
                L1.P1 = SetPoint(X, Y)
            Case 2
                L1.P2 = SetPoint(X, Y)
        End Select
        Redraw
    End If
    lblCurPos.Caption = "X:" & X & "  Y:" & Y
   
End Sub

Private Sub txtod_Change()
    
    Redraw

End Sub

Private Sub chkShow_Click(Index As Integer)
    
    Redraw

End Sub

Private Sub Redraw()
    
    Dim OD          As Single
    Dim OL(1 To 2)  As tLINE

    picCanvas.DrawStyle = vbSolid
    CanvasRedraw picCanvas
    DrawPointRect picCanvas, L1.P1, 1
    DrawPointRect picCanvas, L1.P2, 2
    DrawLine picCanvas, L1
    OD = GetVal(txtod)
    OffsetLine L1, OD, OL
    picCanvas.DrawStyle = vbDot
    If chkShow(0).Value = vbChecked Then DrawLine picCanvas, OL(1)
    If chkShow(1).Value = vbChecked Then DrawLine picCanvas, OL(2)

End Sub
