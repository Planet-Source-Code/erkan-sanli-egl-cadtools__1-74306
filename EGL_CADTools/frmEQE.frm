VERSION 5.00
Begin VB.Form frmEQE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Equation of an Ellipse"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   451
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   574
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCenterX 
      Height          =   285
      Left            =   6600
      TabIndex        =   12
      Text            =   "0"
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtCenterY 
      Height          =   285
      Left            =   6600
      TabIndex        =   11
      Text            =   "0"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtb 
      Height          =   285
      Left            =   6600
      TabIndex        =   8
      Text            =   "3"
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Redraw"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txta 
      Height          =   285
      Left            =   6600
      TabIndex        =   4
      Text            =   "5"
      Top             =   4800
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
   Begin VB.Label Label8 
      Caption         =   "center at"
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
      TabIndex        =   15
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "x"
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
      TabIndex        =   14
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "y"
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
      TabIndex        =   13
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblCurPos 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblCurPos"
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblb 
      Alignment       =   2  'Center
      Caption         =   "b"
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
      TabIndex        =   9
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "b=minor axis"
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
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "a=major axis"
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
      TabIndex        =   6
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "y=+- sqr((ab^2-bx^2)/a^2)"
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
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "(x^2 / a^2) + (y^2 / b^2) = 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   6015
   End
   Begin VB.Label lbla 
      Alignment       =   2  'Center
      Caption         =   "a"
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
      TabIndex        =   1
      Top             =   4800
      Width           =   255
   End
End
Attribute VB_Name = "frmEQE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------
'Equation of an Ellipse
'------------------------

Option Explicit

Private Sub Command1_Click()
    
    Redraw

End Sub

Private Sub Form_Load()
    
    CanvasRescale picCanvas
    Redraw

End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblCurPos.Caption = "X:" & X & "  Y:" & Y
    
End Sub

Private Sub Redraw()
    
    CanvasRedraw picCanvas
    EquationEllipse
        
End Sub

Private Sub EquationEllipse()

'Referance
'http://www.mathopenref.com/coordgeneralellipse.html
    
    Dim X   As Single
    Dim Y   As Single
    Dim A   As Double
    Dim B   As Double
    Dim CX  As Single
    Dim CY  As Single
    
    A = GetVal(txta)
    B = GetVal(txtb)
    CX = GetVal(txtCenterX)
    CY = GetVal(txtCenterY)
    
    For X = -A To A Step 0.01
        Y = Sqr(((A * A * B * B) - (X * X * B * B)) / (A * A))
        picCanvas.PSet (CX + X, CY + -Y)
        picCanvas.PSet (CX + X, CY + Y)
    Next
    
    picCanvas.CurrentX = CX + (A / 2)
    picCanvas.CurrentY = CY + 0.25
    picCanvas.Print "a"
    
    picCanvas.CurrentX = CX - 0.1
    picCanvas.CurrentY = CY + (B / 2)
    picCanvas.Print "b"
    
    picCanvas.DrawWidth = 4
    picCanvas.PSet (CX, CY), vbBlue

End Sub

