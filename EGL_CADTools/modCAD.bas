Attribute VB_Name = "modCAD"
Option Explicit

Private Const sPIDiv180     As Single = 0.0174533 'PI / 180
Private Const ApproachVal   As Double = 0.000001
Public Const Mag            As Single = 0.15 'Point rectangle magnitude

Public Type POINT
    X   As Single
    Y   As Single
End Type

Public Type tLINE
    P1  As POINT
    P2  As POINT
End Type

Public Type tCIRCLE
    C As POINT
    R As Single
End Type

Public Type tELLIPSE
    A As Single
    B As Single
    C As POINT
End Type

Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Function SetPoint(X As Single, Y As Single) As POINT
    
    SetPoint.X = X
    SetPoint.Y = Y

End Function


Public Function SetLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As tLINE
    
    SetLine.P1.X = X1
    SetLine.P1.Y = Y1
    SetLine.P2.X = X2
    SetLine.P2.Y = Y2

End Function

Public Function SetCircle(CX As Single, CY As Single, R As Single) As tCIRCLE
    
    SetCircle.C.X = CX
    SetCircle.C.Y = CY
    SetCircle.R = R

End Function

Public Function SetEllipse(CX As Single, CY As Single, A As Single, B As Single) As tELLIPSE
    
    SetEllipse.C.X = CX
    SetEllipse.C.Y = CY
    SetEllipse.A = A
    SetEllipse.B = B

End Function

Public Sub DrawLine(Canvas As PictureBox, L1 As tLINE)
    
    Canvas.Line (L1.P1.X, L1.P1.Y)-(L1.P2.X, L1.P2.Y)

End Sub

Public Sub DrawCircle(Canvas As PictureBox, C1 As tCIRCLE)
    If C1.R < 0.1 Then C1.R = 0.1
    Canvas.Circle (C1.C.X, C1.C.Y), C1.R

End Sub

Public Sub DrawEllipse(Canvas As PictureBox, E1 As tELLIPSE)
    
    Dim K(1 To 2)   As POINT        'Corner point for API (Ellipse drawing)

    K(1).X = -E1.A * 20 + 200
    K(1).Y = -E1.B * 20 - 200
    K(2).X = E1.A * 20 + 200
    K(2).Y = E1.B * 20 - 200
    Ellipse Canvas.hdc, K(1).X, -K(1).Y, K(2).X, -K(2).Y

End Sub

Public Function Distance(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
    
    Distance = Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2)

End Function

Public Function Div(ByVal Val1 As Single, ByVal Val2 As Single) As Single
    
    If Val2 = 0 Then Val2 = ApproachVal
    Div = CSng(Val1 / Val2)

End Function

Public Sub IntersectionLineLine(L1 As tLINE, L2 As tLINE, I As POINT)

'Referance :
'http://paulbourke.net/geometry/lineline2d/

    Dim M As Single

    M = Div(((L2.P2.X - L2.P1.X) * (L1.P1.Y - L2.P1.Y) - (L2.P2.Y - L2.P1.Y) * (L1.P1.X - L2.P1.X)), _
            ((L2.P2.Y - L2.P1.Y) * (L1.P2.X - L1.P1.X) - (L2.P2.X - L2.P1.X) * (L1.P2.Y - L1.P1.Y)))
    I.X = L1.P1.X + M * (L1.P2.X - L1.P1.X)
    I.Y = L1.P1.Y + M * (L1.P2.Y - L1.P1.Y)
    
End Sub

Public Sub IntersectionLineCircle(L1 As tLINE, C1 As tCIRCLE, F() As Boolean, I() As POINT)

'Referance :
'http://paulbourke.net/geometry/sphereline/
    
    Dim A  As Single
    Dim B  As Single
    Dim C  As Single
    Dim D  As Single
    Dim M  As Single
    
    L1.P1.X = L1.P1.X
    L1.P2.X = L1.P2.X
    C1.C.X = C1.C.X
    L1.P1.Y = L1.P1.Y
    L1.P2.Y = L1.P2.Y
    C1.C.Y = C1.C.Y
    
    A = (L1.P2.X - L1.P1.X) * (L1.P2.X - L1.P1.X) + (L1.P2.Y - L1.P1.Y) * (L1.P2.Y - L1.P1.Y)
    B = 2 * ((L1.P2.X - L1.P1.X) * (L1.P1.X - C1.C.X) + (L1.P2.Y - L1.P1.Y) * (L1.P1.Y - C1.C.Y))
    C = C1.C.X * C1.C.X + C1.C.Y * C1.C.Y + L1.P1.X * L1.P1.X + L1.P1.Y * L1.P1.Y - _
        2 * (C1.C.X * L1.P1.X + C1.C.Y * L1.P1.Y) - C1.R * C1.R
    D = B * B - 4 * A * C
    
    Select Case D
        Case Is < 0
            F(1) = False
            F(2) = False
        Case Is = 0
            F(1) = True
            F(2) = False
            M = -B / (2 * A)
            I(1).X = L1.P1.X + M * (L1.P2.X - L1.P1.X)
            I(1).Y = L1.P1.Y + M * (L1.P2.Y - L1.P1.Y)
        Case Is > 0
            F(1) = True
            F(2) = True
            M = (-B + Sqr(D)) / (2 * A)
            I(1).X = L1.P1.X + M * (L1.P2.X - L1.P1.X)
            I(1).Y = L1.P1.Y + M * (L1.P2.Y - L1.P1.Y)
            M = (-B - Sqr(D)) / (2 * A)
            I(2).X = L1.P1.X + M * (L1.P2.X - L1.P1.X)
            I(2).Y = L1.P1.Y + M * (L1.P2.Y - L1.P1.Y)
    End Select

End Sub

Public Sub IntersectionLineEllipse(L1 As tLINE, E1 As tELLIPSE, F() As Boolean, I() As POINT)

    Dim C  As Single
    Dim D  As Single
    Dim M  As Single
    Dim G(1 To 5) As Single
        
    M = Div((L1.P2.Y - L1.P1.Y), (L1.P2.X - L1.P1.X))
    C = L1.P1.Y - M * L1.P1.X
    D = E1.A * E1.A * M * M + E1.B * E1.B - C * C
    Select Case D
        Case Is < 0
            F(1) = False
            F(2) = False
        Case Is = 0
            F(1) = True
            F(2) = False
            I(1).X = -(E1.A * E1.A * M / C)
            I(1).Y = E1.B * E1.B / C
        Case Is > 0
            F(1) = True
            F(2) = True
            G(1) = -E1.A * E1.A * M * C
            G(2) = E1.A * E1.B
            G(3) = E1.A * E1.A * M * M + E1.B * E1.B - C * C
            If G(3) > 0 Then
                G(3) = Sqr(G(3))
                G(4) = E1.A * E1.A * M * M + E1.B * E1.B
                G(5) = E1.B * E1.B * C
                I(1).X = Div(G(1) + G(2) * G(3), G(4))
                I(1).Y = Div(G(5) + G(2) * M * G(3), G(4))
                I(2).X = Div(G(1) - G(2) * G(3), G(4))
                I(2).Y = Div(G(5) - G(2) * M * G(3), G(4))
            End If
    End Select
    
End Sub

Public Sub IntersectionCircleCircle(C1 As tCIRCLE, C2 As tCIRCLE, F As Boolean, I() As POINT)

'Referance :
'http://paulbourke.net/geometry/2circle/

    Dim A  As Single
    Dim D  As Single
    Dim H  As Single
    Dim K  As POINT
        
    D = Distance(C2.C.X, C2.C.Y, C1.C.X, C1.C.Y)
    
    If D > C1.R + C2.R Then
        'No solution. Circles are seperate.
        F = False
        Exit Sub
    End If
    
    If D < Abs(C1.R - C2.R) Then
        'No solution. One circle is contained within the other.
        F = False
        Exit Sub
    End If
    
    If D = 0 And C1.R = C2.R Then
        'Infinite
        F = False
        Exit Sub
    End If
    
    F = True
    A = (C1.R * C1.R - C2.R * C2.R + D * D) / (2 * D)
    H = Sqr(C1.R * C1.R - A * A)
    K.X = C1.C.X + A * (C2.C.X - C1.C.X) / D
    K.Y = C1.C.Y + A * (C2.C.Y - C1.C.Y) / D
    I(1).X = K.X + H * (C2.C.Y - C1.C.Y) / D
    I(1).Y = K.Y - H * (C2.C.X - C1.C.X) / D
    I(2).X = K.X - H * (C2.C.Y - C1.C.Y) / D
    I(2).Y = K.Y + H * (C2.C.X - C1.C.X) / D

End Sub

Public Sub IntersectionCircleEllipse(C1 As tCIRCLE, E1 As tELLIPSE, F As Boolean, I() As POINT)

    Dim C  As Single
    Dim D  As Single
    Dim M  As Single
    Dim S  As Single
    Dim L  As Single
    Dim K(1 To 5) As Single
    
    
    If Abs(E1.A) < Abs(E1.B) Then
        S = E1.A
        L = E1.B
    Else
        S = E1.B
        L = E1.A
    End If
    If Abs(C1.R) < Abs(L) And Abs(C1.R) > Abs(S) Then
        F = True
        K(1) = C1.R ^ 2 - E1.B ^ 2
        K(2) = E1.A ^ 2 - E1.B ^ 2
        K(3) = Div(K(1), K(2))
        K(4) = E1.A ^ 2 - C1.R ^ 2
        K(5) = Div(K(4), K(2))
        
        If K(3) > 0 Then
            I(1).X = E1.A * Sqr(K(3))
            I(2).X = E1.A * Sqr(K(3))
            I(3).X = -E1.A * Sqr(K(3))
            I(4).X = -E1.A * Sqr(K(3))
        End If
        
        If K(5) > 0 Then
            I(1).Y = E1.B * Sqr(K(5))
            I(2).Y = -E1.B * Sqr(K(5))
            I(3).Y = E1.B * Sqr(K(5))
            I(4).Y = -E1.B * Sqr(K(5))
        End If
    Else
        F = False
    End If

End Sub

Public Sub MidPoint(L1 As tLINE, MP As POINT)

'Referance
'http://www.mathopenref.com/coordmidpoint.html
            
    MP.X = (L1.P1.X + L1.P2.X) / 2
    MP.Y = (L1.P1.Y + L1.P2.Y) / 2
    
End Sub

Public Sub PerpendicularPointOnTheLine(L1 As tLINE, FP As POINT, PP As POINT)

'Referance :
'http://paulbourke.net/geometry/pointline/
    
    Dim u  As Single
    
    u = (((FP.X - L1.P1.X) * (L1.P2.X - L1.P1.X) + (FP.Y - L1.P1.Y) * (L1.P2.Y - L1.P1.Y)) / ((L1.P2.X - L1.P1.X) ^ 2 + (L1.P2.Y - L1.P1.Y) ^ 2))
    PP.X = L1.P1.X + u * (L1.P2.X - L1.P1.X)
    PP.Y = L1.P1.Y + u * (L1.P2.Y - L1.P1.Y)
    
End Sub

Public Sub NearestPointOnTheLine(L1 As tLINE, F As Boolean, FP As POINT, NP As POINT)
    
    Dim D  As Single
    
    PerpendicularPointOnTheLine L1, FP, NP
    D = Distance(NP.X, NP.Y, FP.X, FP.Y)
    F = IIf(Abs(D) < Mag, True, False)      'Mag = approach value

End Sub

Public Sub NearestPointOnTheCircle(C1 As tCIRCLE, F As Boolean, FP As POINT, NP As POINT)
    
    Dim D1  As Single
    Dim D2  As Single
    
    D1 = Distance(C1.C.X, C1.C.Y, FP.X, FP.Y)
    NP.X = C1.C.X + Div((FP.X - C1.C.X), D1) * C1.R
    NP.Y = C1.C.Y + Div((FP.Y - C1.C.Y), D1) * C1.R
    D2 = C1.R - D1
    F = IIf(Abs(D2) < Mag, True, False)      'Mag = approach value

End Sub

Public Sub NearestPointOnTheEllipse(E1 As tELLIPSE, F As Boolean, FP As POINT, NP As POINT)
    
    Dim D   As Single 'Delta Distances
    Dim F1  As Single
    
    F1 = Div(E1.A * E1.B, Sqr(E1.A * E1.A * FP.Y * FP.Y + E1.B * E1.B * FP.X * FP.X))
    NP.X = F1 * FP.X
    NP.Y = F1 * FP.Y
    D = Distance(E1.C.X, E1.C.Y, FP.X, FP.Y) - Distance(E1.C.X, E1.C.Y, NP.X, NP.Y)
    F = IIf(Abs(D) < Mag, True, False)      'Mag = approach value

End Sub

Public Sub TangentCircle(C1 As tCIRCLE, F As Boolean, FP As POINT, TP() As POINT)

'Referance
'http://www.mathopenref.com/coordmidpoint.html
'http://paulbourke.net/geometry/2circle/

    Dim A  As Single
    Dim C  As POINT
    Dim D  As Single
    Dim H  As Single
    Dim K  As POINT

    C.X = (FP.X + C1.C.X) / 2
    C.Y = (FP.Y + C1.C.Y) / 2
    D = Distance(C.X, C.Y, C1.C.X, C1.C.Y)
    If 2 * D > C1.R Then
        F = True
        A = (2 * D * D - C1.R * C1.R) / (2 * D)
        H = Sqr(D * D - A * A)
        K.X = C.X + A * (C1.C.X - C.X) / D
        K.Y = C.Y + A * (C1.C.Y - C.Y) / D
        TP(1).X = K.X + H * (C1.C.Y - C.Y) / D
        TP(1).Y = K.Y - H * (C1.C.X - C.X) / D
        TP(2).X = K.X - H * (C1.C.Y - C.Y) / D
        TP(2).Y = K.Y + H * (C1.C.X - C.X) / D
    Else
        F = False
    End If

End Sub

Public Sub Circle3Points(P() As POINT, C1 As tCIRCLE)

'Referance :
'http://paulbourke.net/geometry/circlefrom3/
    
    Dim MA  As Single
    Dim MB  As Single
    
    If P(1).X = P(2).X And P(2).X = P(3).X Then Exit Sub
    If P(1).Y = P(2).Y And P(2).Y = P(3).Y Then Exit Sub
    
    MA = Div((P(2).Y - P(1).Y), (P(2).X - P(1).X))
    MB = Div((P(3).Y - P(2).Y), (P(3).X - P(2).X))
    
    C1.C.X = Div((MA * MB * (P(1).Y - P(3).Y) + MB * (P(1).X + P(2).X) - MA * (P(2).X + P(3).X)), (2 * (MB - MA)))
    C1.C.Y = Div(-1, MA) * (C1.C.X - ((P(1).X + P(2).X) / 2)) + ((P(1).Y + P(2).Y) / 2)
    C1.R = (Distance(C1.C.X, C1.C.Y, P(1).X, P(1).Y))

End Sub

Public Sub CircleLLR(L1 As tLINE, L2 As tLINE, R As Single, C1 As tCIRCLE)
    
    Dim OL1(1 To 2) As tLINE   '1st Line Offset points
    Dim OL2(1 To 2) As tLINE   '2nd line Offset points
    Dim E           As POINT   'lines points average
    Dim OM(1 To 4)  As POINT   'Offset line mid point
    Dim D(1 To 4)   As Single  'Distances between om and e
    Dim IL(1 To 2)  As tLINE   'Offset Lines for Intersection
    
    OffsetLine L1, R, OL1
    OffsetLine L2, R, OL2
    E.X = (L1.P1.X + L1.P2.X + L2.P1.X + L2.P2.X) / 4
    E.Y = (L1.P1.Y + L1.P2.Y + L2.P1.Y + L2.P2.Y) / 4
    MidPoint OL1(1), OM(1)
    MidPoint OL1(2), OM(2)
    MidPoint OL2(1), OM(3)
    MidPoint OL2(2), OM(4)
    D(1) = Distance(E.X, E.Y, OM(1).X, OM(1).Y)
    D(2) = Distance(E.X, E.Y, OM(2).X, OM(2).Y)
    D(3) = Distance(E.X, E.Y, OM(3).X, OM(3).Y)
    D(4) = Distance(E.X, E.Y, OM(4).X, OM(4).Y)
    If D(1) < D(2) Then
        IL(1) = OL1(1)
    Else
        IL(1) = OL1(2)
    End If
    If D(3) < D(4) Then
        IL(2) = OL2(1)
    Else
        IL(2) = OL2(2)
    End If
    IntersectionLineLine IL(1), IL(2), C1.C
    C1.R = R

End Sub

Public Sub CircleLCR(L1 As tLINE, C1 As tCIRCLE, R As Single, F() As Boolean, DC() As tCIRCLE)
    
    Dim OL1(1 To 2) As tLINE    'Line Offset
    Dim OC1         As tCIRCLE  'Circle Offset
    Dim OM(1 To 2)  As POINT    'Offset line mid point
    Dim D(1 To 2)   As Single   'Distances between om and center point
    Dim IL          As tLINE    'Offset Line for Intersection
    Dim I(1 To 2)   As POINT
    
    OffsetLine L1, R, OL1
    MidPoint OL1(1), OM(1)
    MidPoint OL1(2), OM(2)
    D(1) = Distance(C1.C.X, C1.C.Y, OM(1).X, OM(1).Y)
    D(2) = Distance(C1.C.X, C1.C.Y, OM(2).X, OM(2).Y)
    If D(1) < D(2) Then
        IL = OL1(1)
    Else
        IL = OL1(2)
    End If
    OC1.C = C1.C
    OC1.R = C1.R + R
    IntersectionLineCircle IL, OC1, F, I
    If F(1) Then DC(1).C = I(1): DC(1).R = R
    If F(2) Then DC(2).C = I(2): DC(2).R = R
    
End Sub

Public Sub CircleCCR(C1 As tCIRCLE, C2 As tCIRCLE, R As Single, F As Boolean, DC() As tCIRCLE)
    
    Dim OC1         As tCIRCLE  'Circle Offset
    Dim OC2         As tCIRCLE  'Circle Offset
    Dim I(1 To 2)   As POINT

    OC1.C = C1.C
    OC1.R = C1.R + R
    OC2.C = C2.C
    OC2.R = C2.R + R
    IntersectionCircleCircle OC1, OC2, F, I
    If F Then
        DC(1).C = I(1): DC(1).R = R
        DC(2).C = I(2): DC(2).R = R
    End If

End Sub

Public Sub CircleLLL(L1 As tLINE, L2 As tLINE, L3 As tLINE, DC As tCIRCLE)
    
    Dim bs12 As tLINE   'Bisector 1-2
    Dim bs13 As tLINE   'Bisector 1-3
    Dim FP As POINT
    Dim PP As POINT
    
    Bisector2Lines L1, L2, bs12
    Bisector2Lines L2, L3, bs13
    IntersectionLineLine bs12, bs13, FP
    PerpendicularPointOnTheLine L3, FP, PP
    DC.C = FP
    DC.R = Distance(FP.X, FP.Y, PP.X, PP.Y)
    
End Sub

Public Sub DivideLine(L1 As tLINE, N As Long, DP() As POINT)
    
    Dim u       As Single
    Dim V       As Single
    Dim idxD    As Long
    
    u = Div((L1.P2.X - L1.P1.X), N)
    V = Div((L1.P2.Y - L1.P1.Y), N)
    ReDim DP(1 To N)
    For idxD = 1 To N
        DP(idxD).X = L1.P1.X + u * idxD
        DP(idxD).Y = L1.P1.Y + V * idxD
    Next

End Sub

Public Sub DivideCircle(C1 As tCIRCLE, N As Long, DP() As POINT)
    
    Dim A       As Double
    Dim idxD    As Long
    
    A = (360 / N) * sPIDiv180
    ReDim DP(1 To N)
    For idxD = 1 To N
        DP(idxD).X = C1.C.X + Cos(A * idxD) * C1.R
        DP(idxD).Y = C1.C.Y + Sin(A * idxD) * C1.R
    Next

End Sub

Public Sub OffsetLine(L1 As tLINE, OD As Single, OL() As tLINE)
    
'Referance
'Please see Line equation (frmEQL, Extra)

    Dim DeltaX  As Single
    Dim DeltaY  As Single
    Dim LenLine As Single
    Dim OP      As POINT
    
    DeltaX = L1.P2.X - L1.P1.X
    DeltaY = L1.P2.Y - L1.P1.Y
    LenLine = Distance(L1.P1.X, L1.P1.Y, L1.P2.X, L1.P2.Y)
    OP = SetPoint(Div(DeltaY, LenLine) * OD, Div(DeltaX, LenLine) * OD)
    OL(1) = SetLine(L1.P1.X - OP.X, L1.P1.Y + OP.Y, L1.P2.X - OP.X, L1.P2.Y + OP.Y)
    OL(2) = SetLine(L1.P1.X + OP.X, L1.P1.Y - OP.Y, L1.P2.X + OP.X, L1.P2.Y - OP.Y)

End Sub

Private Sub Bisector2Lines(L1 As tLINE, L2 As tLINE, BL As tLINE)
    
    Dim LenLine As Single
    Dim I1  As POINT
    Dim I2  As POINT
    Dim BS As POINT
    
    IntersectionLineLine L1, L2, BS
    
    LenLine = Distance(L1.P1.X, L1.P1.Y, L1.P2.X, L1.P2.Y)
    I1.X = BS.X + Div(L1.P2.X - L1.P1.X, LenLine)
    I1.Y = BS.Y + Div(L1.P2.Y - L1.P1.Y, LenLine)
    
    LenLine = Distance(L2.P1.X, L2.P1.Y, L2.P2.X, L2.P2.Y)
    I2.X = BS.X + Div(L2.P2.X - L2.P1.X, LenLine)
    I2.Y = BS.Y + Div(L2.P2.Y - L2.P1.Y, LenLine)
    BL = SetLine(BS.X, BS.Y, (I1.X + I2.X) / 2, (I1.Y + I2.Y) / 2)

End Sub
'
'Public Function Slope(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
'
'    Slope = Div((Y2 - Y1), (X2 - X1))
'
'End Function

