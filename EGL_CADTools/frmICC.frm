VERSION 5.00
Begin VB.Form frmICC 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Intersection of two circles"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   418
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   572
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
Attribute VB_Name = "frmICC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------
'Intersection of two circles
'------------------------

Option Explicit

Dim idx         As Byte     'Selected point index
Dim C1          As tCIRCLE
Dim C2          As tCIRCLE
Dim HR1         As POINT    'radius handel points
Dim HR2         As POINT

Private Sub Form_Load()
    
    C1 = SetCircle(4, 3, 4)
    C2 = SetCircle(-3, -1, 5)
    HR1 = SetPoint(C1.C.X + C1.R, C1.C.Y)
    HR2 = SetPoint(C2.C.X + C2.R, C2.C.Y)
    CanvasRescale picCanvas
    Redraw

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        idx = 0
        If GrabPoint(X, Y, C1.C) Then idx = 1
        If GrabPoint(X, Y, HR1) Then idx = 2
        If GrabPoint(X, Y, C2.C) Then idx = 3
        If GrabPoint(X, Y, HR2) Then idx = 4
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
            Case 3
                C2.C = SetPoint(X, Y)
                HR2 = SetPoint(X + C2.R, Y)
            Case 4
                HR2 = SetPoint(X, Y)
                C2.R = Distance(C2.C.X, C2.C.Y, HR2.X, HR2.Y)
        End Select
        Redraw
    End If
    lblCurPos.Caption = "X:" & X & "  Y:" & Y

End Sub

Private Sub Redraw()
    
    Dim F  As Boolean        'Intersection Flag
    Dim IP(1 To 2)  As POINT 'Intersection points

    CanvasRedraw picCanvas
    DrawPointRect picCanvas, C1.C, 1
    DrawPointRect picCanvas, HR1, 2
    DrawPointRect picCanvas, C2.C, 3
    DrawPointRect picCanvas, HR2, 4
    DrawCircle picCanvas, C1
    DrawCircle picCanvas, C2
    IntersectionCircleCircle C1, C2, F, IP
    If F Then
        picCanvas.Circle (IP(1).X, IP(1).Y), Mag
        picCanvas.Circle (IP(2).X, IP(2).Y), Mag
    End If
    
End Sub
