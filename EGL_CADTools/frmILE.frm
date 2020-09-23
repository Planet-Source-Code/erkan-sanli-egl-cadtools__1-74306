VERSION 5.00
Begin VB.Form frmILE 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Intersection of a Line and an Ellipse"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   417
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
   Begin VB.Label Label1 
      Caption         =   "Ellipse center at (0,0)"
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   840
      Width           =   2055
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
Attribute VB_Name = "frmILE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------
'Intersection of a Line and an Ellipse
'------------------------

Option Explicit


Dim idx         As Byte         'Selected point index
Dim L1          As tLINE
Dim E1          As tELLIPSE
Dim HA          As POINT        'Major axis handel
Dim HB          As POINT        'Minor axis handel

Private Sub Form_Load()
    
    L1 = SetLine(-5, -5, 5, 3)
    E1 = SetEllipse(0, 0, 5, 3)
    HA = SetPoint(E1.C.X + E1.A, E1.C.Y)
    HB = SetPoint(E1.C.X, E1.C.Y + E1.B)
    CanvasRescale picCanvas
    Redraw

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        idx = 0
        If GrabPoint(X, Y, L1.P1) Then idx = 1
        If GrabPoint(X, Y, L1.P2) Then idx = 2
        If GrabPoint(X, Y, HA) Then idx = 3
        If GrabPoint(X, Y, HB) Then idx = 4
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
                HA.X = X
                E1.A = X
            Case 4
                HB.Y = Y
                E1.B = Y
        End Select
        Redraw
    End If
    lblCurPos.Caption = "X:" & X & "  Y:" & Y

End Sub

Private Sub Redraw()
    
    Dim F(1 To 2) As Boolean  'Intersection flags
    Dim IP(1 To 2) As POINT   'Intersection points

    CanvasRedraw picCanvas
    DrawPointRect picCanvas, L1.P1, 1
    DrawPointRect picCanvas, L1.P2, 2
    DrawPointRect picCanvas, HA, 3
    DrawPointRect picCanvas, HB, 4
    DrawLine picCanvas, L1
    DrawEllipse picCanvas, E1
    IntersectionLineEllipse L1, E1, F, IP
    If F(1) Then picCanvas.Circle (IP(1).X, IP(1).Y), Mag
    If F(2) Then picCanvas.Circle (IP(2).X, IP(2).Y), Mag

End Sub
