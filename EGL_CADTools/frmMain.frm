VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAD Tools by Erkan"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "CLLL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   5640
      TabIndex        =   52
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "TTT - 3 lines tangents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   600
      TabIndex        =   51
      Top             =   6120
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "CCCR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   5640
      TabIndex        =   50
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "TTR - 2 circles tangents and radius"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   600
      TabIndex        =   49
      Top             =   5880
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Miscellaneous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   48
      Top             =   6480
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "CLCR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   5640
      TabIndex        =   47
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "TTR - A line and a circle tangents and radius"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   600
      TabIndex        =   46
      Top             =   5640
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Education Version"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   45
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "CLLR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   5640
      TabIndex        =   44
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "TTR - 2 lines tangents and radius"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   600
      TabIndex        =   43
      Top             =   5400
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "MOFL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   5640
      TabIndex        =   42
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Offset lines (Parallel lines)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   600
      TabIndex        =   41
      Top             =   7200
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Divide the circle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   600
      TabIndex        =   40
      Top             =   6960
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "MDVC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   5640
      TabIndex        =   39
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "MDVL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   5640
      TabIndex        =   38
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Divide the line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   600
      TabIndex        =   37
      Top             =   6720
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "STC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   5640
      TabIndex        =   36
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Tangent lines of a circle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   600
      TabIndex        =   35
      Top             =   4560
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "EQQ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   34
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Quadratic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   33
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "C3P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   5640
      TabIndex        =   32
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "3 points"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   600
      TabIndex        =   31
      Top             =   5160
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Draw a circle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   30
      Top             =   4920
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "SPL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   5640
      TabIndex        =   29
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Perpendicular point of a line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   600
      TabIndex        =   28
      Top             =   4320
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "SNE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   5640
      TabIndex        =   27
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Nearest point of an ellipse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   600
      TabIndex        =   26
      Top             =   4080
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "SNC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   5640
      TabIndex        =   25
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "SNL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   5640
      TabIndex        =   24
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "SMP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   23
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Special Points of Geometry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   22
      Top             =   3120
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "ICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   5640
      TabIndex        =   21
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "ICC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   5640
      TabIndex        =   20
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "ILE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5640
      TabIndex        =   19
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "ILC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5640
      TabIndex        =   18
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "ILL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5640
      TabIndex        =   17
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Intersections"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "EQE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   15
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "EQL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   14
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Equations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "EQC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   12
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Nearest point of a circle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   600
      TabIndex        =   11
      Top             =   3840
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Nearest point of the line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   600
      TabIndex        =   10
      Top             =   3600
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Mid point of the line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   600
      TabIndex        =   9
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Circle Circle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Line Ellipse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   6
      Top             =   2280
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Line Circle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   4
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Line Line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Ellipse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Circle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Circle Ellipse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   600
      TabIndex        =   0
      Top             =   2760
      Width           =   4935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click(Index As Integer)
    
    Select Case Index
        Case 0:  Unload Me: End
        Case 1:  frmEQL.Show 'Equation of a Line
        Case 2:  frmEQC.Show 'Equation of a Circle
        Case 3:  frmEQE.Show 'Equation of a Ellipse
        Case 4:  frmEQQ.Show 'Equation of a Quadratic
        
        Case 5:  frmILL.Show 'Intersection Line Line
        Case 6:  frmILC.Show 'Intersection Line Circle
        Case 7:  frmILE.Show 'Intersection Line Ellipse
        Case 8:  frmICC.Show 'Intersection Circle Circle
        Case 9:  frmICE.Show 'Intersection Circle Ellipse

        Case 10: frmSMP.Show 'Midpoint of the line
        Case 11: frmSNL.Show 'Nearest point on the line
        Case 12: frmSNC.Show 'Nearest point on the circle
        Case 13: frmSNE.Show 'Nearest point on the ellipse
        Case 14: frmSPL.Show 'Perpendicular point of a line
        Case 15: frmSTC.Show 'Tangent points of a circle
        
        Case 16: frmC3P.Show  'Draw a circle from 3 Points
        Case 17: frmCLLR.Show 'Draw a circle from 2 lines and radius
        Case 18: frmCLCR.Show 'Draw a circle from a line, a circle and radius
        Case 19: frmCCCR.Show 'Draw a circle from 2 circles and radius
        Case 20: frmCLLL.Show 'Draw a circle from 3 lines
        
        Case 21: frmMDVL.Show 'Divide the line
        Case 22: frmMDVC.Show 'Divide the circle
        Case 23: frmMOFL.Show 'Offset Line
    
    End Select

End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    PaintBlack
    Label1(Index).BackColor = vbBlue
    Label1(Index).ForeColor = vbWhite

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    PaintBlack

End Sub

Private Sub PaintBlack()
    
    Dim idx As Byte
    
    For idx = 0 To 23
        Label1(idx).BackColor = vbWhite
        Label1(idx).ForeColor = vbBlack
    Next

End Sub
