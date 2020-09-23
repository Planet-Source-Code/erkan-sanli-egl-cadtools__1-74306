VERSION 5.00
Begin VB.Form frmICE 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Intersection of a Circle and an Ellipse"
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
   Begin VB.Label Label2 
      Caption         =   "A circle center at (0,0)"
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "An ellipse center at (0,0)"
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   2400
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
Attribute VB_Name = "frmICE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------
'Intersection of a Circle and an Ellipse
'------------------------

Option Explicit

Dim idx         As Byte         'Selected point index
Dim C1          As tCIRCLE
Dim E1          As tELLIPSE
Dim HR          As POINT
Dim HA          As POINT
Dim HB          As POINT

Private Sub Form_Load()
    
    C1 = SetCircle(0, 0, 4)
    E1 = SetEllipse(0, 0, 6, 3)
    HR = SetPoint(C1.C.X + C1.R, C1.C.Y)
    HA = SetPoint(E1.C.X + E1.A, E1.C.Y)
    HB = SetPoint(E1.C.X, E1.C.Y + E1.B)
    CanvasRescale picCanvas
    Redraw

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        idx = 0
        If GrabPoint(X, Y, HR) Then idx = 1
        If GrabPoint(X, Y, HA) Then idx = 2
        If GrabPoint(X, Y, HB) Then idx = 3
    End If
 
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        Select Case idx
            Case 1
                HR = SetPoint(X, Y)
                C1.R = Distance(C1.C.X, C1.C.Y, HR.X, HR.Y)
            Case 2
                HA.X = Abs(X)
                E1.A = HA.X
            Case 3
                HB.Y = Abs(Y)
                E1.B = HB.Y
        End Select
        Redraw
    End If
    lblCurPos.Caption = "X:" & X & "  Y:" & Y

End Sub

Private Sub Redraw()
    
    Dim F As Boolean         'Intersection flag
    Dim IP(1 To 4)  As POINT 'Intersection points

    CanvasRedraw picCanvas
    DrawPointRect picCanvas, HR, 1
    DrawPointRect picCanvas, HA, 2
    DrawPointRect picCanvas, HB, 3
    DrawCircle picCanvas, C1
    DrawEllipse picCanvas, E1
    IntersectionCircleEllipse C1, E1, F, IP
    If F Then
        picCanvas.Circle (IP(1).X, IP(1).Y), Mag
        picCanvas.Circle (IP(2).X, IP(2).Y), Mag
        picCanvas.Circle (IP(3).X, IP(3).Y), Mag
        picCanvas.Circle (IP(4).X, IP(4).Y), Mag
    End If

End Sub
