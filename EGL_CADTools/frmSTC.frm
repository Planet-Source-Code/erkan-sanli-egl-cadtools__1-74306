VERSION 5.00
Begin VB.Form frmSTC 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tangent points of a Circle"
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
   Begin VB.CheckBox chkShow 
      Caption         =   "2. Tangent Line"
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   3
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "1. Tangent Line"
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   2
      Top             =   5400
      Width           =   1695
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
Attribute VB_Name = "frmSTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------
'Tangent points of a Circle
'------------------------

Option Explicit

Dim idx         As Byte     'Selected point index
Dim C1          As tCIRCLE
Dim HR          As POINT    'Handel radius
Dim FP          As POINT    'Float point


Private Sub Form_Load()
    
    C1 = SetCircle(3, 1, 3)
    HR = SetPoint(C1.C.X + C1.R, C1.C.Y)
    FP = SetPoint(-5, -5)
    chkShow(0).Value = vbChecked
    CanvasRescale picCanvas
    Redraw

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        idx = 0
        If GrabPoint(X, Y, FP) Then idx = 1
        If GrabPoint(X, Y, C1.C) Then idx = 2
        If GrabPoint(X, Y, HR) Then idx = 3
    End If
 
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        Select Case idx
            Case 1
                FP = SetPoint(X, Y)
            Case 2
                C1.C = SetPoint(X, Y)
                HR = SetPoint(X + C1.R, Y)
            Case 3
                HR = SetPoint(X, Y)
                C1.R = Distance(C1.C.X, C1.C.Y, HR.X, HR.Y)
        End Select
        Redraw
    End If
    lblCurPos.Caption = "X:" & X & "  Y:" & Y

End Sub

Private Sub chkShow_Click(Index As Integer)
    
    Redraw

End Sub

Private Sub Redraw()

    Dim F As Boolean         'Tangent Flag
    Dim TP(1 To 2)  As POINT 'Tangent points
    
    picCanvas.DrawStyle = vbSolid
    CanvasRedraw picCanvas
    DrawPointRect picCanvas, FP, 1
    DrawPointRect picCanvas, C1.C, 2
    DrawPointRect picCanvas, HR, 3
    DrawCircle picCanvas, C1
    TangentCircle C1, F, FP, TP
    If F Then
        If chkShow(0).Value = vbChecked Then picCanvas.Circle (TP(1).X, TP(1).Y), Mag
        If chkShow(1).Value = vbChecked Then picCanvas.Circle (TP(2).X, TP(2).Y), Mag
        picCanvas.DrawStyle = vbDot
        If chkShow(0).Value = vbChecked Then picCanvas.Line (FP.X, FP.Y)-(TP(1).X, TP(1).Y)
        If chkShow(1).Value = vbChecked Then picCanvas.Line (FP.X, FP.Y)-(TP(2).X, TP(2).Y)
    End If
    
End Sub
