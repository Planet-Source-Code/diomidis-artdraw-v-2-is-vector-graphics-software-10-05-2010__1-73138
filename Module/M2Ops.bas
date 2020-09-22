Attribute VB_Name = "M2Ops"
' Routines for manipulating 2-dimensional
' vectors and matrices.
Type Coord
     X As Single
     Y As Single
End Type

Option Explicit

' Create a 2-dimensional identity matrix.
Public Sub m2Identity(m() As Single)
Dim I As Integer
Dim j As Integer

    For I = 1 To 3
        For j = 1 To 3
            If I = j Then
                m(I, j) = 1
            Else
                m(I, j) = 0
            End If
        Next j
    Next I
End Sub

' Create a translation matrix for translation by
' distances tx and ty.
Public Sub m2Translate(Result() As Single, ByVal tx As Single, ByVal ty As Single)
    m2Identity Result
    Result(3, 1) = tx
    Result(3, 2) = ty
End Sub

' Create a scaling matrix for scaling by factors
' of sx and sy.
Public Sub m2Scale(Result() As Single, ByVal sx As Single, ByVal sy As Single)
    m2Identity Result
    Result(1, 1) = sx
    Result(2, 2) = sy
End Sub

' Create a Skew matrix for Skew by factors
' of sx and sy.
Public Sub m2Skew(Result() As Single, ByVal sx As Single, ByVal sy As Single)
    m2Identity Result
    Result(1, 2) = 1 - sy
    Result(2, 1) = 1 - sx
End Sub

' Create a rotation matrix for rotating by the
' given angle (in radians).
Public Sub m2Rotate(Result() As Single, ByVal theta As Single)
    m2Identity Result
    Result(1, 1) = Cos(theta)
    Result(1, 2) = Sin(theta)
    Result(2, 1) = -Result(1, 2)
    Result(2, 2) = Result(1, 1)
End Sub

' Create a rotation matrix that rotates the point
' (x, y) onto the X axis.
Public Sub m2RotateIntoX(Result() As Single, ByVal X As Single, ByVal Y As Single)
Dim d As Single

    m2Identity Result
    d = Sqr(X * X + Y * Y)
    Result(1, 1) = X / d
    Result(1, 2) = -Y / d
    Result(2, 1) = -Result(1, 2)
    Result(2, 2) = Result(1, 1)
End Sub

' Create a scaling matrix for scaling by factors
' of sx and sy at the point (x, y).
Public Sub m2ScaleAt(Result() As Single, _
                     ByVal sx As Single, ByVal sy As Single, _
                     ByVal X As Single, ByVal Y As Single)
                    
Dim t(1 To 3, 1 To 3) As Single
Dim s(1 To 3, 1 To 3) As Single
Dim T_Inv(1 To 3, 1 To 3) As Single
Dim m(1 To 3, 1 To 3) As Single

    ' Translate the point to the origin.
    m2Translate t, -X, -Y

    ' Compute the inverse translation.
    m2Translate T_Inv, X, Y

    ' Scale.
    m2Scale s, sx, sy

    ' Combine the transformations.
    m2MatMultiply m, t, s
    m2MatMultiply Result, m, T_Inv
End Sub

' Create a matrix for reflecting across the line
' passing through (x, y) in direction <dx, dy>.
Public Sub m2ReflectAcross(Result() As Single, ByVal X As Single, ByVal Y As Single, ByVal dX As Single, ByVal dy As Single)
Dim t(1 To 3, 1 To 3) As Single
Dim R(1 To 3, 1 To 3) As Single
Dim s(1 To 3, 1 To 3) As Single
Dim R_Inv(1 To 3, 1 To 3) As Single
Dim T_Inv(1 To 3, 1 To 3) As Single
Dim m1(1 To 3, 1 To 3) As Single
Dim m2(1 To 3, 1 To 3) As Single
Dim M3(1 To 3, 1 To 3) As Single

    ' Translate the point to the origin.
    m2Translate t, -X, -Y

    ' Compute the inverse translation.
    m2Translate T_Inv, X, Y

    ' Rotate so the direction vector lies in the Y axis.
    m2RotateIntoX R, dX, dy

    ' Compute the inverse translation.
    m2RotateIntoX R_Inv, dX, -dy

    ' Reflect across the X axis.
    m2Scale s, 1, -1

    ' Combine the transformations.
    m2MatMultiply m1, t, R     ' T * R
    m2MatMultiply m2, s, R_Inv ' S * R_Inv
    m2MatMultiply M3, m1, m2   ' T * R * S * R_Inv

    ' T * R * S * R_Inv * T_Inv
    m2MatMultiply Result, M3, T_Inv
End Sub

' Create a Skew matrix
' of sx and sy at the point (x, y).
Public Sub m2SkewAt(Result() As Single, ByVal sx As Single, ByVal sy As Single, ByVal X As Single, ByVal Y As Single)
                    
Dim t(1 To 3, 1 To 3) As Single
Dim s(1 To 3, 1 To 3) As Single
Dim T_Inv(1 To 3, 1 To 3) As Single
Dim m(1 To 3, 1 To 3) As Single

    ' Translate the point to the origin.
    m2Translate t, -X, -Y

    ' Compute the inverse translation.
    m2Translate T_Inv, X, Y

    ' Skew.
    m2Skew s, sx, sy

    ' Combine the transformations.
    m2MatMultiply m, t, s           ' T * S
    m2MatMultiply Result, m, T_Inv  ' T * S * T_Inv
End Sub

' Create a rotation matrix for rotating through
' angle theta around the point (x, y).
Public Sub m2RotateAround(Result() As Single, ByVal theta As Single, ByVal X As Single, ByVal Y As Single)
Dim t(1 To 3, 1 To 3) As Single
Dim R(1 To 3, 1 To 3) As Single
Dim T_Inv(1 To 3, 1 To 3) As Single
Dim m(1 To 3, 1 To 3) As Single

    ' Translate the point to the origin.
    m2Translate t, -X, -Y

    ' Compute the inverse translation.
    m2Translate T_Inv, X, Y

    ' Rotate.
    m2Rotate R, theta

    ' Combine the transformations.
    m2MatMultiply m, t, R
    m2MatMultiply Result, m, T_Inv
End Sub

' Multiply a point and a matrix.
Public Sub m2PointMultiply(ByRef X As Single, ByRef Y As Single, a() As Single)
Dim NewX As Single
Dim NewY As Single

    NewX = X * a(1, 1) + Y * a(2, 1) + a(3, 1)
    NewY = X * a(1, 2) + Y * a(2, 2) + a(3, 2)
    X = NewX
    Y = NewY
End Sub
' Set copy = orig.
Public Sub m2PointCopy(copy() As Single, orig() As Single)
Dim I As Integer

    For I = 1 To 3
        copy(I) = orig(I)
    Next I
End Sub

' Set copy = orig.
Public Sub m2MatCopy(copy() As Single, orig() As Single)
Dim I As Integer
Dim j As Integer

    For I = 1 To 3
        For j = 1 To 3
            copy(I, j) = orig(I, j)
        Next j
    Next I
End Sub

' Apply a transformation matrix to a point.
Public Sub m2Apply(Result() As Single, v() As Single, a() As Single)
    Result(1) = v(1) * a(1, 1) + v(2) * a(2, 1) + a(3, 1)
    Result(2) = v(1) * a(1, 2) + v(2) * a(2, 2) + a(3, 2)
    Result(3) = 1#
End Sub

' Multiply two transformation matrices.
Public Sub m2MatMultiply(Result() As Single, a() As Single, b() As Single)
    Result(1, 1) = a(1, 1) * b(1, 1) + a(1, 2) * b(2, 1)
    Result(1, 2) = a(1, 1) * b(1, 2) + a(1, 2) * b(2, 2)
    Result(1, 3) = 0#
    Result(2, 1) = a(2, 1) * b(1, 1) + a(2, 2) * b(2, 1)
    Result(2, 2) = a(2, 1) * b(1, 2) + a(2, 2) * b(2, 2)
    Result(2, 3) = 0#
    Result(3, 1) = a(3, 1) * b(1, 1) + a(3, 2) * b(2, 1) + b(3, 1)
    Result(3, 2) = a(3, 1) * b(1, 2) + a(3, 2) * b(2, 2) + b(3, 2)
    Result(3, 3) = 1#
End Sub

Public Function m2Atn2(ByVal Y As Single, ByVal X As Single) As Single
   If X = 0 Then
      m2Atn2 = IIf(Y = 0, Pi / 4, Sgn(Y) * Pi / 2)
   Else
      m2Atn2 = Atn(Y / X) + (1 - Sgn(X)) * Pi / 2
   End If
End Function

'Public Function m2GetAngle3P(P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI) As Single
Public Function m2GetAngle3P(X1 As Single, Y1 As Single, _
                             X2 As Single, Y2 As Single, _
                             X3 As Single, Y3 As Single) As Single
Debug.Print X1, Y1, X2, Y2, X3, Y3
' Retreive the angle formed by this 3 points
' P1<---->P2 e P1<----->P3
' rather than GetAngle3P this function doesn't need of a
' parallel edge and it returns the real internal angle close to the P1 point

      ' / P2
     ' /
    ' /
   ' /
  ' /
'P1 \ <-a°
   ' \
    ' \
     ' \
      ' \ P3
'

Dim I As Integer, An As Single
Dim Alfa As Single
Dim a As Double, b As Double, c As Double, PS As Double
Dim Ds1 As Double, Ds2 As Double, Ds3 As Double
Dim Q1 As Coord, Q2 As Coord, Q3 As Coord
Dim P1 As Coord, P2 As Coord, P3 As Coord

   P1.X = X1
   P1.Y = Y1
   P2.X = X2
   P2.Y = Y2
   P3.X = X3
   P3.Y = Y3
   
Const Rg# = 200 / Pi
    Ds1 = Dist(P1.X, P1.Y, P2.X, P2.Y)
    Ds2 = Dist(P1.X, P1.Y, P3.X, P3.Y)
    Ds3 = Dist(P3.X, P3.Y, P2.X, P2.Y)

    a = Ds3
    b = Ds1
    c = Ds2

    If a = 0 Or b = 0 Or c = 0 Then Exit Function

    PS = (a + b + c) * 0.5
    If PS < c Then GoTo ErrorAngle
    If PS < a Or PS < b Then GoTo ErrorAngle

    On Error Resume Next
    Alfa = 2 * Atn(((PS - b) * (PS - c) / PS / (PS - a)) ^ 0.5) * Rg#

    Alfa = m2An360(Alfa)

    Q1 = P1
    Q2 = P2
    Q3.X = P2.X
    Q3.Y = P1.Y

    An = GetAngle3P(Q1, Q2, Q3)
    If An <> 0 Then
        Q3.X = P3.X - P1.X
        Q3.Y = P3.Y - P1.Y
        Q3 = m2RotatePoint(Q3.X, Q3.Y, -An)
        Q3.X = Q3.X + P1.X
        Q3.Y = Q3.Y + P1.Y
    End If

    If Q3.Y < Q1.Y Then Alfa = 360 - Alfa
    m2GetAngle3P = Alfa

Exit Function

ErrorAngle:

    m2GetAngle3P = 0

End Function

Private Function GetAngle3P(P1 As Coord, P2 As Coord, P3 As Coord) As Single
'Public Function m2GetAngle3P(X1 As Single, Y1 As Single, _
                             X2 As Single, Y2 As Single, _
                             X3 As Single, Y3 As Single) As Single
' Calculate angle from edges
' P1<---->P2 e P1<----->P3

' Note:
' It returns the angle 0-360 referred by the edge P1-P3 always parallel to the X axe
' if that edge (P1-P3) is not parallel the function will wrong the result value
'
' Next checks in wich square P2 is contained
' to set the relative angle (0-90 , 91-180, 181-270 or 271,360)
'

Dim I As Integer, k As Integer, m As Integer
Dim X1 As Double, Y1 As Double
Dim X2 As Double, Y2 As Double
Dim Alfa As Single
Dim a As Double, b As Double, c As Double, PS As Double
Dim Fd As Boolean
Dim Q1 As Coord, Q2 As Coord
Dim Ds1 As Single, Ds2 As Single, Ds3 As Single

Const Rg# = 200 / Pi

X1 = P1.X
Y1 = P1.Y

X2 = P2.X
Y2 = P2.Y

Ds1 = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)

X2 = P3.X
Y2 = P3.Y

Ds2 = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)

X1 = P2.X
Y1 = P2.Y

Ds3 = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)

a = Ds3
b = Ds1
c = Ds2

If a = 0 Or b = 0 Or c = 0 Then GoTo Parallel
PS = (a + b + c) * 0.5
If PS < c Then GoTo ErrorAngle
If PS < a Or PS < b Then GoTo ErrorAngle

On Error Resume Next
Alfa = 2 * Atn(((PS - b) * (PS - c) / PS / (PS - a)) ^ 0.5) * Rg#

' Alfa now is in centesimal units (0-400) need to convert it with An360
Alfa = m2An360(Alfa)

' Check Sqares

Parallel:

X1 = P1.X
Y1 = P1.Y

X2 = P2.X
Y2 = P2.Y


If X1 = X2 Then
    If Y1 > Y2 Then
        Alfa = 270
    ElseIf Y1 < Y2 Then
        Alfa = 90
    End If
ElseIf Y1 = Y2 Then
    If X1 > X2 Then
        Alfa = 180
    ElseIf X1 < X2 Then
        Alfa = 0
    End If
ElseIf X1 > X2 And Y1 < Y2 Then ' II°
    Alfa = 90 - Alfa + 90
ElseIf X1 > X2 And Y1 > Y2 Then ' III°
    Alfa = Alfa + 180
ElseIf X1 < X2 And Y1 > Y2 Then ' IV°
    Alfa = 90 - Alfa + 270
End If

GetAngle3P = Alfa

Exit Function

ErrorAngle:
    GetAngle3P = 0
End Function

Public Function m2An360(An As Single) As Single
' Transform an Angle from Centesimal 0,400 to
' 0, 360
If An <> 0 Then
   m2An360 = An / 1.11111111111111
Else
   m2An360 = 0
End If

End Function

Function Dist(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
        Dist = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
End Function

' Rotate a single Point using Rad function to converts Degree to Radians
Private Function m2RotatePoint(X As Single, Y As Single, Angle As Single) As Coord
Dim XA As Single, YA As Single
Dim mSin As Single, mCos As Single
Dim P As Coord
   P.X = X
   P.Y = Y
If Angle <> 0 Then
   mSin = Sin(Angle * Pi / 180): mCos = Cos(Angle * Pi / 180)
   XA = mCos * P.X - mSin * P.Y
   YA = mSin * P.X + mCos * P.Y
   m2RotatePoint.X = XA
   m2RotatePoint.Y = YA
Else
   m2RotatePoint = P
End If

End Function
