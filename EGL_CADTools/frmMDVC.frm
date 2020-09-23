VERSION 5.00
Begin VB.Form frmMDVC 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Divide the Circle"
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
   Begin VB.TextBox txtn 
      Height          =   285
      Left            =   6600
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
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      Caption         =   "n"
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
      TabIndex        =   4
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Division number"
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   5280
      Width           =   1695
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
Attribute VB_Name = "frmMDVC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------
'Divide the Circle
'------------------------

Option Explicit


Dim idx         As Byte     'Selected point index
Dim C1          As tCIRCLE
Dim HR1         As POINT    'Radius handel points

Private Sub Form_Load()
    
    C1 = SetCircle(2, 1, 3)
    HR1 = SetPoint(C1.C.X + C1.R, C1.C.Y)
    CanvasRescale picCanvas
    Redraw

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        idx = 0
        If GrabPoint(X, Y, C1.C) Then idx = 1
        If GrabPoint(X, Y, HR1) Then idx = 2
    End If
 
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        Select Case idx
            Case 1
                C1.C = SetPoint(X, Y)
                HR1 = SetPoint(X + C1.R, Y)
            Case 2
                HR1 = SetPoint(X, Y)
                C1.R = Distance(C1.C.X, C1.C.Y, HR1.X, HR1.Y)
        End Select
        Redraw
    End If
    lblCurPos.Caption = "X:" & X & "  Y:" & Y

End Sub

Private Sub txtn_Change()
    
    Redraw

End Sub

Private Sub Redraw()
    
    Dim DP()    As POINT    'Divide points
    Dim N       As Long
    Dim idxD    As Long
    
    CanvasRedraw picCanvas
    DrawPointRect picCanvas, C1.C, 1
    DrawPointRect picCanvas, HR1, 2
    DrawCircle picCanvas, C1
    N = GetVal(txtn)
    DivideCircle C1, N, DP
    For idxD = 1 To UBound(DP)
        picCanvas.Circle (DP(idxD).X, DP(idxD).Y), Mag
    Next
    
End Sub
