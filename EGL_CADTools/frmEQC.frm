VERSION 5.00
Begin VB.Form frmEQC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Equation of a Circle"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   578
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCenterY 
      Height          =   285
      Left            =   6600
      TabIndex        =   11
      Text            =   "0"
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox txtCenterX 
      Height          =   285
      Left            =   6600
      TabIndex        =   9
      Text            =   "0"
      Top             =   4440
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
   Begin VB.TextBox txtr 
      Height          =   285
      Left            =   6600
      TabIndex        =   4
      Text            =   "5"
      Top             =   5160
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
      TabIndex        =   12
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label5 
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
      TabIndex        =   10
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblCurPos 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblCurPos"
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      TabIndex        =   7
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "r=radius"
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
      Caption         =   "y=+- sqr(r^2-x^2)"
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
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "a^2 + b^2 = c^2"
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
   Begin VB.Label lblr 
      Alignment       =   2  'Center
      Caption         =   "r"
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
      Top             =   5160
      Width           =   255
   End
End
Attribute VB_Name = "frmEQC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------
'Equation of a Circle
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
    EquationCircle
        
End Sub

Private Sub EquationCircle()

'Referance
'http://www.mathopenref.com/coordbasiccircle.html
    
    Dim X   As Single
    Dim Y   As Single
    Dim R   As Double
    Dim CX  As Single
    Dim CY  As Single
    
    R = GetVal(txtr)
    CX = GetVal(txtCenterX)
    CY = GetVal(txtCenterY)
    
    For X = -R To R Step 0.01
        Y = Sqr(R * R - X * X)
        picCanvas.PSet (CX + X, CY + Y)
        picCanvas.PSet (CX + X, CY + -Y)
    Next
    
    picCanvas.DrawWidth = 4
    picCanvas.PSet (CX, CY), vbBlue

End Sub
