Attribute VB_Name = "modCommon"
Option Explicit

Public Function GrabPoint(X As Single, Y As Single, P As POINT) As Boolean

    If X > P.X - Mag And _
       X < P.X + Mag And _
       Y > P.Y - Mag And _
       Y < P.Y + Mag Then
       GrabPoint = True
    End If

End Function

Public Sub DrawPointRect(obj As Object, P As POINT, idx As Integer)
    
    obj.Line (P.X - Mag, P.Y - Mag)-(P.X + Mag, P.Y + Mag), , B
    obj.CurrentX = P.X + Mag
    obj.CurrentY = P.Y - Mag
    obj.Print "P" & idx

End Sub

Public Function GetVal(txt As TextBox) As Single
    
    txt.Text = Replace(txt.Text, ".", ",")
    If IsNumeric(txt.Text) Then
        GetVal = CSng(txt.Text)
    Else
        GetVal = 0
        txt.Text = 0
    End If
    
End Function

Public Sub CanvasRedraw(picBox As PictureBox)
    
    Dim idx As Integer
    
    picBox.Cls
    picBox.DrawWidth = 1
    
    'Grid
    For idx = picBox.ScaleLeft To picBox.ScaleWidth
        picBox.Line (idx, picBox.ScaleTop)-(idx, picBox.ScaleHeight), RGB(230, 230, 230)
        picBox.ForeColor = RGB(180, 180, 180)
        picBox.CurrentX = idx
        picBox.CurrentY = 0
        picBox.Print idx
    Next
    
    For idx = picBox.ScaleTop To picBox.ScaleHeight Step -1
        picBox.Line (picBox.ScaleLeft, idx)-(picBox.ScaleWidth, idx), RGB(230, 230, 230)
        picBox.ForeColor = RGB(180, 180, 180)
        picBox.CurrentX = 0
        picBox.CurrentY = idx
        picBox.Print idx
    Next
    
    'Crosshairs
    picBox.Line (picBox.ScaleLeft, 0)-(picBox.ScaleWidth, 0), RGB(250, 150, 150)
    picBox.Line (0, picBox.ScaleTop)-(0, picBox.ScaleHeight), RGB(150, 250, 150)
    picBox.ForeColor = RGB(0, 0, 0)

End Sub

Public Sub CanvasRescale(picBox As PictureBox)
    
    picBox.ScaleLeft = -10
    picBox.ScaleTop = 10
    picBox.ScaleWidth = 20
    picBox.ScaleHeight = -20

End Sub

