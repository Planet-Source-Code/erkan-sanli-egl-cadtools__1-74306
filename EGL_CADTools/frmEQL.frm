VERSION 5.00
Begin VB.Form frmEQL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Equation of a Line"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkExtra 
      Caption         =   "Extra"
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Redraw"
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtb 
      Height          =   285
      Left            =   6600
      TabIndex        =   7
      Text            =   "2"
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtm 
      Height          =   285
      Left            =   6600
      TabIndex        =   6
      Text            =   "2"
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
   Begin VB.Label lblEx 
      BackColor       =   &H00FFFF00&
      Caption         =   "[m<>0]"
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   14
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblEx 
      BackColor       =   &H00FFFF00&
      Caption         =   "Perpendicular of a line"
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   13
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblEx 
      BackColor       =   &H00FFFF00&
      Caption         =   "-1 / m= -deltax / deltay"
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
      Index           =   2
      Left            =   6360
      TabIndex        =   12
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblCurPos 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblCurPos"
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "m= deltay / deltax"
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
      Left            =   6360
      TabIndex        =   10
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "blue point= (0,b)"
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
      Left            =   6360
      TabIndex        =   9
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "b= intercept"
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
      Left            =   6360
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "m= slope of the line"
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
      Left            =   6360
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "y=mx+b"
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
      TabIndex        =   3
      Top             =   6240
      Width           =   4095
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
      TabIndex        =   2
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label lblm 
      Alignment       =   2  'Center
      Caption         =   "m"
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
Attribute VB_Name = "frmEQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------
'Equation of a Line
'------------------------

Option Explicit

Private Sub chkExtra_Click()
    
    lblEx(0).Visible = chkExtra.Value
    lblEx(1).Visible = chkExtra.Value
    lblEx(2).Visible = chkExtra.Value
    Redraw

End Sub

Private Sub Command1_Click()
    
    Redraw

End Sub

Private Sub Form_Load()
    
    chkExtra_Click
    CanvasRescale picCanvas
    Redraw

End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblCurPos.Caption = "X:" & X & "  Y:" & Y
    
End Sub

Private Sub Redraw()
    
    CanvasRedraw picCanvas
    EquationLine
        
End Sub

Private Sub EquationLine()

    Dim X As Single
    Dim Y As Single
    Dim M As Double
    Dim B As Double
        
    M = GetVal(txtm)
    B = GetVal(txtb)
    
    For X = picCanvas.ScaleLeft To -picCanvas.ScaleLeft Step 0.01
        Y = M * X + B
        picCanvas.PSet (X, Y)
    Next
    
    If M <> 0 And chkExtra.Value Then
        For X = picCanvas.ScaleLeft To -picCanvas.ScaleLeft Step 0.01
            Y = (-1 / M) * X + B
            picCanvas.PSet (X, Y), vbCyan
        Next
    End If
        
    picCanvas.DrawWidth = 4
    picCanvas.PSet (0, B), vbBlue

End Sub
