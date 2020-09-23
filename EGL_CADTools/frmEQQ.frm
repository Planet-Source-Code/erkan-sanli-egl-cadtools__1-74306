VERSION 5.00
Begin VB.Form frmEQQ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Equation of Quadratic"
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
   Begin VB.TextBox txtc 
      Height          =   285
      Left            =   6600
      TabIndex        =   7
      Text            =   "0"
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtb 
      Height          =   285
      Left            =   6600
      TabIndex        =   5
      Text            =   "0"
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Redraw"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txta 
      Height          =   285
      Left            =   6600
      TabIndex        =   3
      Text            =   "0,1"
      Top             =   4440
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
   Begin VB.Label lblCurPos 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblCurPos"
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "c"
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
      TabIndex        =   8
      Top             =   5160
      Width           =   255
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
      TabIndex        =   6
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "y = ax^2 + bx + c"
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
      Top             =   4440
      Width           =   255
   End
End
Attribute VB_Name = "frmEQQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------
'Equation of Quadratic
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
    EquationQuadratic
        
End Sub

Private Sub EquationQuadratic()

'Referance
'http://www.mathopenref.com/quadraticexplorer.html
    
    Dim X As Single
    Dim Y As Single
    Dim A As Double
    Dim B As Double
    Dim C As Double
    
    A = GetVal(txta)
    B = GetVal(txtb)
    C = GetVal(txtc)
    
    For X = picCanvas.ScaleLeft To -picCanvas.ScaleLeft Step 0.01
        Y = (A * X * X) + (B * X) + C
        picCanvas.PSet (X, Y)
    Next
    

End Sub
