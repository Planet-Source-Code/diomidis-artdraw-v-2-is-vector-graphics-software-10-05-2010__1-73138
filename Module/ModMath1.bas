Attribute VB_Name = "ModMath1"
Const PI As Double = 3.14159265358979

Public Type Coord
     X As Single
     Y As Single
End Type
Public Type Box
     Min As POINTAPI
     Max As POINTAPI
End Type

Public Function Angle(ByVal n1 As Double, ByVal n2 As Double) As Double
'You might not want to worry about how this function works,
'unless you are familiar with the trigonometry.  Just pay
'attention to the function structure and try to understand
'how the function operates, other than the math.
'I chose a long function to show how convenient functions
'can be.  If you need to calculate the angle that a point
'is from 0,0 in 15 different places, it's much better to
'have 15 function calls than to have this code all over.
'If you realize you did it wrong, you would have to fix it
'in all fifteen places instead of just inside the function

' We'll use this to store our result
Dim Result As Double

'We have to check if n1 is zero, because we will
'get a divide by zero error later if it is.  In
'the event that it is zero, we will send out the
'proper value and then quit the function
If n1 = 0 Then
    If n2 > 0 Then
        Angle = 90
        Exit Function
    ElseIf n2 < 0 Then
        Angle = 270
        Exit Function
    Else
        'this only happens when both values are zero.
        'since the function is designed to find the angle
        'between (0, 0) and the input set, it is impossible
        'to get an angle if both points are at the same
        'place, so we just return...it will just give
        'us a zero.
        Exit Function
    End If
End If

'mathy things.  We are taking the arc-tangent
Result = Atn(n2 / n1)
'and converting it to degrees
Result = Result * 180 / 3.14159265358979

'checking if it the first value (X) is less than
'zero...if it is, we need to adjust the results.
If n1 < 0 Then Result = Result + 180

'Now make sure it stays between 0 and 360 degrees
If Result < 0 Then Result = Result + 360
If Result > 359 Then Result = Result - 360

'return the answer
Angle = Result
End Function

Public Function AreaTrapezoidParallelogram(Base1 As Double, Base2 As Double, Altitude As Double) As Double
    'calculate area of Trapazoid, Parallelagram

    Dim Temp As Double
    Temp = Base1 + Base2
    Temp = Temp * Altitude
    Temp = Temp * 0.5
    AreaTrapezoidParallelogram = Temp

End Function

Public Function AreaTriangleRight(Base As Double, Altitude As Double) As Double
       AreaTriangleRight = (Base * Altitude) * 0.5
End Function

Public Function Atn2(X As Double, Y As Double) As Single
    Const NearZero = 0.000000001
    If Y = 0 Then Y = NearZero
    Atn2 = (Atn(Abs(X) / Abs(Y)) * Sgn(X) - 3.141592 / 2) * Sgn(Y)
End Function

'Calculate Arc Cosecant
'********************************************************
'* Name : Arccos
'* Description :
'********************************************************
Public Function ArcCos(ByVal sRadians As Single) As Single
    Dim Appo As Single
    Appo = -sRadians * sRadians + 1
    If Appo <= 0 Then
        ArcCos = IIf(sRadians > 0, 0, 3.1415)
    Else
        ArcCos = Atn(-sRadians / Sqr(Appo)) + 2 * Atn(1)
    End If
End Function

'Public Function GetInternalAngle3P(p1 as PointApi, p2 as PointApi, P3 as PointApi) As Single
Public Function GetInternalAngle3P(ByVal X1 As Single, ByVal Y1 As Single, _
                                   ByVal X2 As Single, ByVal Y2 As Single, _
                                   ByVal X3 As Single, ByVal Y3 As Single) As Single
    
    Dim P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI
    
    P1.X = X1
    P1.Y = Y1
    P2.X = X2
    P2.Y = Y2
    P3.X = X3
    P3.Y = Y3
    
    ' Retreive the angle formed by this 3 points
    ' P1<---->P2 e P1<----->P3
    ' rather than GetAngle3P this function doesn't need of a parallel edge
    ' and it returns the real internal angle close to the P1 point
    '       / P3
    '      /
    '     /
    '    /
    '   /
    'P1 \ <-a°
    '    \
    '     \
    '      \
    '       \ P2

    Dim i As Integer
    Dim Alfa As Double
    Dim a As Double, B As Double, C As Double, PS As Double
    Dim Ds1 As Double, Ds2 As Double, Ds3 As Double
    Dim Q1 As POINTAPI, Q2 As POINTAPI, Q3 As POINTAPI

    Const Rg# = 200 / PI
    Ds1 = Dist(P1.X, P1.Y, P2.X, P2.Y)
    Ds2 = Dist(P1.X, P1.Y, P3.X, P3.Y)
    Ds3 = Dist(P3.X, P3.Y, P2.X, P2.Y)

    a = Ds3
    B = Ds1
    C = Ds2

    If a = 0 Or B = 0 Or C = 0 Then Exit Function

    PS = (a + B + C) * 0.5
    If PS < C Then GoTo ErroreAngolo
    If PS < a Or PS < B Then GoTo ErroreAngolo

    On Error Resume Next
    Alfa = 2 * Atn(((PS - B) * (PS - C) / PS / (PS - a)) ^ 0.5) * Rg#
    Alfa = An360(Alfa)

    Q1 = P1
    Q2 = P2
    Q3.X = P2.X
    Q3.Y = P1.Y

    An! = GetAngle3P(Q1.X, Q1.Y, Q2.X, Q2.Y, Q3.X, Q3.Y)

    If An! <> 0 Then
        Q3.X = P3.X - P1.X
        Q3.Y = P3.Y - P1.Y
        Q3 = Rotate(Q3.X, Q3.Y, -An!)
        Q3.X = Q3.X + P1.X
        Q3.Y = Q3.Y + P1.Y
    End If

    If Q3.Y < Q1.Y Then Alfa = 360 - Alfa
    If Format(Alfa, "0") = 90 Then
      GetInternalAngle3P = 90
    ElseIf Format(Alfa, "0") = 180 Then
      GetInternalAngle3P = 180
    ElseIf Format(Alfa, "0") = 270 Then
      GetInternalAngle3P = 270
    ElseIf Format(Alfa, "0") = 0 Then
      GetInternalAngle3P = 0
    ElseIf Format(Alfa, "0") = 360 Then
       GetInternalAngle3P = 360
    Else
       GetInternalAngle3P = Alfa
    End If
    
    Exit Function
ErroreAngolo:
    GetInternalAngle3P = 0
End Function

'Convert a centesimal angle (topographic) to cartesian angle
Function An360(An As Double) As Double
' Transform an Angle from Centesimal 0,400 to
' 0, 360
    If An <> 0 Then
        An360 = An / 1.11111111111111
    Else
        An360 = 0
    End If

End Function

'Calculate the Angle given by three points with a parallel edge
'Function GetAngle3P(P1 as PointApi, P2 as PointApi, P3 as PointApi) As Single
Function GetAngle3P(ByVal pX1 As Single, ByVal pY1 As Single, _
                    ByVal pX2 As Single, ByVal pY2 As Single, _
                    ByVal px3 As Single, ByVal py3 As Single) As Single
' Calculate angle from edges
' P1<---->P2 e P1<----->P3
' Note:
' It returns the angle 0-360 referred by the edge P1-P3 always parallel to the X axe
' if that edge (P1-P3) is not parallel the function will wrong the result value
'
' Next checks in wich square P2 is contained
' to set the relative angle (0-90 , 91-180, 181-270 or 271,360)

    Dim i As Integer, k As Integer, m As Integer
    Dim X1 As Double, Y1 As Double
    Dim X2 As Double, Y2 As Double
    Dim Alfa As Double
    Dim a As Double, B As Double, C As Double, PS As Double
    Dim Fd As Boolean
    Dim Q1 As POINTAPI, Q2 As POINTAPI
    Dim P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI
    P1.X = pX1
    P1.Y = pY1
    P2.X = pX2
    P2.Y = pY2
    P3.X = px3
    P3.Y = py3
    
    Const Rg# = 200 / PI

    X1 = P1.X
    Y1 = P1.Y

    X2 = P2.X
    Y2 = P2.Y

    Ds1# = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)

    X2 = P3.X
    Y2 = P3.Y

    Ds2# = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)

    X1 = P2.X
    Y1 = P2.Y

    Ds3# = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)

    a = Ds3#
    B = Ds1#
    C = Ds2#

    If a = 0 Or B = 0 Or C = 0 Then GoTo Parallel

    PS = (a + B + C) * 0.5
    If PS < C Then GoTo ErroreAngolo
    If PS < a Or PS < B Then GoTo ErroreAngolo

    On Error Resume Next
    Alfa = 2 * Atn(((PS - B) * (PS - C) / PS / (PS - a)) ^ 0.5) * Rg#
    ' Alfa now is in centesimal units (0-400) need to convert it with An360
    Alfa = An360(Alfa)
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
ErroreAngolo:
    GetAngle3P = 0
End Function


'Find Angle of an arbitrary point in the 2D space
'********************************************************
'* Name : FindAngle
'* Description : Getting the angle formed by origin 0,0 and the given point x,y
'********************************************************
Public Function FindAngle(ByVal X As Single, ByVal Y As Single) As Single
    '
    If X <> 0 Then
        FindAngle = Degrees(Atn(Y / X))
    Else
        FindAngle = Sgn(Y) * 90
    End If
    '
    If X < 0 Then
        FindAngle = FindAngle + 180
    Else
        If Y < 0 Then
            FindAngle = FindAngle + 360
        End If
    End If
    '
End Function

'Converting Degrees To Radians
Function Rad(Angle As Single) As Double
    Rad = Angle * (PI / 180)
End Function

'Public Static Function PI() As Double
'PI = Atn(90) * 4
'End Function

'Converting Radians To Degrees
Function Degrees(xRad As Single) As Double
    Degrees = xRad * 180 / PI
End Function

'Rotating a Point
'Public Function Rotate(p As POINTAPI, Angle As Single) As POINTAPI
Public Function Rotate(ByVal X As Single, ByVal Y As Single, ByVal Angle As Single) As POINTAPI
' Rotate a single Point using Rad function to converts Degree to Radians
   
    Dim XA As Double, YA As Double
    Dim Seno As Double, Coseno As Double
    Dim P As POINTAPI
    P.X = X
    P.Y = Y
    If Angle <> 0 Then
        Seno = Sin(Rad(Angle)): Coseno = Cos(Rad(Angle))
        XA = Coseno * P.X - Seno * P.Y
        YA = Seno * P.X + Coseno * P.Y
        Rotate.X = Format(XA, "0.00")
        Rotate.Y = Format(YA, "0.00")
        'x = XA
        'y = YA
    Else
        Rotate = P
    End If
End Function

''Circle passing between three points
'Sub CircumCircle(X1 As Double, y1 As Double, _
'                 x2 As Double, y2 As Double, _
'                 x3 As Double, y3 As Double, _
'                 xLin As Double, yLin As Double, radius As Double)
'    ' Calculate the elements of a circle passing for three points
'    ' x1,y1 x2,y2 x3,y3
'    ' the point coord will be stored in the variables
'    ' xlin1 and ylin1
'    ' and the radius
'
'    Dim cec_a As Double, cec_b As Double, cec_c As Double
'    Dim cec_d As Double, cec_e As Double, cec_f As Double
'    Dim cec_g As Double, cec_h As Double, cec_i As Double
'
'    cec_a = X1 - x2
'    cec_b = y1 - y2
'    cec_c = x2 - x3
'    cec_d = y2 - y3
'    cec_e = X1 * X1 - x2 * x2
'    cec_f = y1 * y1 - y2 * y2
'    cec_g = x2 * x2 - x3 * x3
'    cec_h = y2 * y2 - y3 * y3
'    cec_i = 2 * (cec_a * cec_d - cec_b * cec_c)
'
'    If cec_i <> 0 Then
'        ' centre
'        xlin1 = (cec_d * (cec_e + cec_f) - cec_b * (cec_g + cec_h)) / cec_i
'        ylin1 = (cec_a * (cec_g + cec_h) - cec_c * (cec_e + cec_f)) / cec_i
'        ' Radius
'        radius = Sqr((iv(1, 1) - xlin1) ^ 2 + (iv(1, 2) - ylin1) ^ 2)
'    Else
'        zlin1 = 0
'    End If
'
'End Sub
'
'Comparing Boxes
Public Function CompareBoxes(Box1 As Box, Box2 As Box) As Integer
' Checks the relation between two boxes
' user type Box:
'     Min as PointApi
'     Max as PointApi
' Return value:
'  0 - Box1 external to Box 2
'  1 - Box2 withinl Box 1
'  2 - Box1 and Box intersect

    Dim C As Boolean

    C = Box2.Max.X < Box1.Min.X Or Box2.Max.Y < Box1.Min.Y Or _
    Box2.Min.X > Box1.Max.X Or Box2.Min.Y > Box1.Max.Y

' c=True Box1 outside box2

    If C Then
        CompareBoxes = 0 ' External
    Else
        C = Box2.Max.X < Box1.Max.X And Box2.Max.Y < Box1.Max.Y And _
        Box2.Min.X > Box1.Min.X Or Box2.Min.Y > Box1.Min.Y
        If C Then
            CompareBoxes = 1 ' Box2 inside Box1
        Else
            CompareBoxes = 2 ' intersection found
        End If
    End If

End Function

'Point distance from a given line
Function PointDistanceFromLine(pX As Double, pY As Double, _
                               X1 As Double, Y1 As Double, _
                               X2 As Double, Y2 As Double) As Double

' Return the distance of a point from a given line
' x1,y1 First Vertex
' x2,y2 Second Vertex
' x0,y0 Point

    Dim dX As Double, dy As Double

    If (X1 = X2) Then
        PointDistanceFromLine = Abs(X1 - pX)
    ElseIf Y1 = Y2 Then
        PointDistanceFromLine = Abs(Y1 - pY)
    Else
        dX = X2 - X1
        dy = Y2 - Y1
        PointDistanceFromLine = Abs(dy * pX - dX * pY + X2 * Y1 - X1 * Y2) / Sqr(dX * dX + dy * dy)
    End If
End Function

'Calculate Distance between two points
'Public Function Dist(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
'         Dist = Format(Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2), "0.00")
'End Function
Function DistTrig(a As Single, B As Single, C As Single, _
                  a1 As Single, b1 As Single, C1 As Single) As Single
         
         Dim SelectType As Integer
         If a = 0 And B <> 0 And C <> 0 Then
            SelectType = 1
         ElseIf a <> 0 And B = 0 And C <> 0 Then
            SelectType = 2
         ElseIf a <> 0 And B <> 0 And C = 0 Then
            SelectType = 3
         Else
            Exit Function
         End If
         
         Select Case SelectType
         Case 1
              DistTrig = Sqr(B * B + C * C - 2 * B * C * Cos(Rad(a1)))
         Case 2
              DistTrig = Sqr(C * C + a * a - 2 * C * a * Cos(Rad(b1)))
         Case 3
              DistTrig = Sqr(a * a + B * B - 2 * a * B * Cos(Rad(C1)))
         End Select
         
End Function

Public Sub Azimut(dX As Double, dy As Double, az As Double)
    ' calcola l'angolo azimutale
    If dy = 0 Then
    If dX > 0 Then az = PI / 2
    If dX < 0 Then az = 3 * PI / 2
    If dX = 0 Then az = 999
        Exit Sub
    End If
    az = Atn(dX / dy)
    If dy < 0 Then az = az + PI
    If az < 0 Then az = az + 2 * PI
End Sub

Function Intersection(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single)

    Dim de As Double, dn As Double, Lt As Double, az As Double

    de = X2 - X1 'Val(in_eb.Text) - Val(in_ea.Text)
    dn = Y2 - Y1 'Val(in_nb.Text) - Val(in_na.Text)

    Call Azimut(de, dn, az)
    Lt = Sqr(de * de + dn * dn)
    Dim a As Double, ab As Double
    a = Lt
    ab = az

    ' distanza tra le due stazioni
    Dim sg As Double
    sg = PI - Val(in_a.Text) * PI / 200 - Val(in_b.Text) * PI / 200 - Val(in_c.Text) * PI / 200
    Dim m As Double, v As Double, n As Double, B As Double
    m = Sin(Val(in_c.Text) * PI / 200) / Sin(sg)
    v = Sin(Val(in_c.Text) * PI / 200 + Val(in_d.Text) * PI / 200)
    n = v / Sin(Val(in_b.Text) * PI / 200 + Val(in_c.Text) * PI / 200 + Val(in_d.Text) * PI / 200)
    v = Sqr(m * m + n * n - 2 * m * n * Cos(Val(in_a.Text) * PI / 200))
    B = a / v

    'risoluzione triangoli
    Dim C As Double, d As Double, e As Double, X As Double, fi As Double
    C = B * m
    d = B * n
    X = (a * a + C * C - d * d) / 2 / a / C
    If Abs(X) = 1 Then GoTo salto
    fi = PI / 2 - Atn(X / Sqr(1 - X * X))
    GoTo salto3

salto:
    fi = PI * (1 - X) / 2
salto3:
    fi = fi - sg
    e = B * Sin(Val(in_a.Text) * PI / 200 + Val(in_b.Text) * PI / 200) / Sin(sg)

    ' direzioni
    Dim aq As Double, AP As Double
    aq = ab + fi
    AP = aq + sg

    'coordinate
    ep = Val(in_ea.Text) + C * Sin(AP)
    np = Val(in_na.Text) + C * Cos(AP)
    eq = Val(in_ea.Text) + e * Sin(aq)
    nq = Val(in_na.Text) + e * Cos(aq)

End Function

Public Function FindDraw(X1 As Single, Y1 As Single, _
                         X2 As Single, Y2 As Single, _
                         SelectX As Single, SelectY As Single, _
                         Optional Snap As Double = 5) As Boolean

'Dim X1 As Single, Y1 As Integer, X2 As Integer, Y2 As Integer, Snap As Integer
Dim Sine As Double, X As Single, Y As Single

'                                /| b(x2, y2)
'                               / |
'                              /  |
'                             /   |
'                            /    |
'                 a(x1, y1) /_____| c

'Label1.Caption = "X=" & X & ", Y=" & Y
'X1 = Text1(0)
'Y1 = Text1(1)
'X2 = Text1(2)
'Y2 = Text1(3)
'Snap = Text1(4)
Sine = (Y1 - Y2) / (X2 - X1) ' bc / ac

If Int((X - X1) * Sine) < (Y1 - Y + Snap) And _
   Int((X - X1) * Sine) > (Y1 - Y - Snap) And _
   X < X2 And X > X1 Then
      'Picture1.MousePointer = 2
      FindDraw = True
Else
      'Picture1.MousePointer = 0
      FindDraw = False
End If

End Function

Function Perpend2(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single, Dis As Single) As POINTAPI

' The line P1-P2 has P3 adjacent. The resulted point will be the next vertex
' of a line perpendicular to line P1-P2 with first vertex P3 and with Lenght = Dis
Dim P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI
Dim Angolo As Single
Dim dX As Double, dy As Double
Dim P As POINTAPI
Dim Q1 As POINTAPI, Q2 As POINTAPI, Q3 As POINTAPI

' Translate line to the origin
P1.X = X1
P1.Y = Y1
P2.X = X2
P2.Y = Y2
P3.X = X3
P3.Y = Y3

dX = Min(P1.X, P2.X)
dy = Min(P1.Y, P2.Y)

Q1.X = P1.X - dX
Q1.Y = P1.Y - dy
Q2.X = P2.X - dX
Q2.Y = P2.Y - dy

' Calculate the slope

Q3.X = Q2.X
Q3.Y = Q1.Y

Angolo = GetAngle3P(Q1.X, Q1.Y, Q2.X, Q2.Y, Q3.X, Q3.Y)

Q3.X = P3.X - dX
Q3.Y = P3.Y - dy

' Rotate the point (on origin)

Q3 = Rotate(Q3.X, Q3.Y, -Angolo)

P.X = Q3.X
P.Y = Q3.Y - Dis
P = Rotate(P.X, P.Y, Angolo)
P.X = P.X + dX
P.Y = P.Y + dy

Perpend2 = P

End Function

Function Max(ByVal X1 As Single, ByVal X2 As Single) As Single
     If X1 >= X2 Then
        Max = X1
     Else
        Max = X2
     End If
End Function


Function Min(ByVal X1 As Single, ByVal X2 As Single) As Single
     If X1 < X2 Then
        Min = X1
     Else
        Min = X2
     End If
End Function

'Õðïëïãéóìüò ðáñÜëëçëçò
Function GiveOffset(Patima As Single, _
                    X1 As Single, Y1 As Single, _
                    X2 As Single, Y2 As Single, _
                    outP() As POINTAPI)
         
         Dim P() As POINTAPI
         Dim Ang1 As Single, Ang2 As Single, Ang3 As Single, Ang4 As Single
         Dim mX1 As Single, mY1 As Single, mX2 As Single, mY2 As Single
         Dim mX3 As Single, mY3 As Single, mX4 As Single, mY4 As Single
         ReDim outP(1 To 2) As POINTAPI
'         Ang1 = GetInternalAngle3P(X1, Y1, X4, Y4, X2, Y2) / 2 'ok
'         Ang2 = GetInternalAngle3P(X2, Y2, X1, Y1, X3, Y3) / 2 '
         ''Debug.Print "Ang1:" + Str(Ang1) + ", Ang2:" + Str(Ang2)
         GoSub NewPoint
         GivePointPlane mX1, mY1, mX2, mY2, Ang1, Ang2, 0, P, , , , Patima
         outP(1).X = P(1).X
         outP(2).X = P(2).X
         outP(1).Y = P(1).Y
         outP(2).Y = P(2).Y
         
Exit Function
NewPoint:
    mX1 = X1
    mX2 = X2
    mY1 = Y1
    mY2 = Y2
Return
End Function

'Õðïëïãéóìüò åðéöÜíåéáò ìå áíåîáñôçôç ôçí êáèå ðëåõñÜ
Function GivePlane2(P1 As Single, P2 As Single, P3 As Single, P4 As Single, _
                    X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                    X3 As Single, Y3 As Single, X4 As Single, Y4 As Single, _
                    cp1 As Single, cp2 As Single, cp3 As Single, cP4 As Single, _
                    outP() As POINTAPI)
         
         Dim pn1() As POINTAPI, pn2() As POINTAPI, pn3() As POINTAPI, pn4() As POINTAPI
         Dim Ang1 As Single, Ang2 As Single, Ang3 As Single, Ang4 As Single
         Dim mX1 As Single, mY1 As Single, mX2 As Single, mY2 As Single
         Dim mX3 As Single, mY3 As Single, mX4 As Single, mY4 As Single
         ReDim outP(1 To 4) As POINTAPI
'         Y1 = -Y1
'         Y2 = -Y2
'         Y3 = -Y3
'         Y4 = -Y4
         Ang1 = GetInternalAngle3P(X1, Y1, X2, Y2, X4, Y4) / 2
         Ang2 = GetInternalAngle3P(X2, Y2, X3, Y3, X1, Y1) / 2
         Ang3 = GetInternalAngle3P(X3, Y3, X4, Y4, X2, Y2) / 2
         Ang4 = GetInternalAngle3P(X4, Y4, X1, Y1, X3, Y3) / 2
         ''Debug.Print "Ang1:" + Str(Ang1) + ", Ang2:" + Str(Ang2) + ", Ang3:" + Str(Ang3) + ", Ang4:" + Str(Ang3)
         GoSub NewPoint
         GivePoint mX1, mY1, mX2, mY2, Ang1, Ang2, P1, pn1, , , , -cp1
         GoSub NewPoint
         GivePoint mX2, mY2, mX3, mY3, Ang2, Ang3, P2, pn2, , , , -cp2
         GoSub NewPoint
         GivePoint mX3, mY3, mX4, mY4, Ang3, Ang4, P3, pn3, , , , -cp3
          GoSub NewPoint
         GivePoint mX4, mY4, mX1, mY1, Ang4, Ang1, P4, pn4, , , , -cP4
         outP(1).X = pn1(4).X
         outP(2).X = pn2(4).X
         outP(1).Y = pn1(4).Y
         outP(2).Y = pn2(4).Y
         
         outP(3).X = pn3(4).X
         outP(3).Y = pn3(4).Y
         outP(4).X = pn4(4).X
         outP(4).Y = pn4(4).Y
         ''Debug.Print Str(outP(1).X) + "," + Str(outP(1).Y) + "," + Str(outP(2).X) + "," + Str(outP(2).Y) + "," + Str(outP(3).X) + "," + Str(outP(3).Y) + "," + Str(outP(4).X) + "," + Str(outP(4).Y)
         ''Debug.Print "Dist1:", Dist(outP(1).X, outP(1).Y, outP(2).X, outP(2).Y)
         ''Debug.Print "Dist2:", Dist(outP(2).X, outP(2).Y, outP(3).X, outP(3).Y)
         ''Debug.Print "Dist3:", Dist(outP(3).X, outP(3).Y, outP(4).X, outP(4).Y)
         ''Debug.Print "Dist4:", Dist(outP(4).X, outP(4).Y, outP(1).X, outP(1).Y)
Exit Function

NewPoint:
    mX1 = X1
    mX2 = X2
    mX3 = X3
    mX4 = X4
    mY1 = Y1
    mY2 = Y2
    mY3 = Y3
    mY4 = Y4
Return
End Function

'Õðïëïãéóìüò åðéöÜíåéáò
Function GivePlane(Patima As Single, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                   X3 As Single, Y3 As Single, X4 As Single, Y4 As Single, outP() As POINTAPI)
         
         Dim P() As POINTAPI
         Dim Ang1 As Single, Ang2 As Single, Ang3 As Single, Ang4 As Single
         Dim mX1 As Single, mY1 As Single, mX2 As Single, mY2 As Single
         Dim mX3 As Single, mY3 As Single, mX4 As Single, mY4 As Single
         ReDim outP(1 To 4) As POINTAPI
         Ang1 = GetInternalAngle3P(X1, Y1, X2, Y2, X4, Y4) ' / 2 'ok
         Ang2 = GetInternalAngle3P(X2, Y2, X3, Y3, X1, Y1) '/ 2 '
         Ang3 = GetInternalAngle3P(X3, Y3, X4, Y4, X2, Y2) ' / 2 '
         Ang4 = GetInternalAngle3P(X4, Y4, X1, Y1, X3, Y3) '/ 2 '
         ''Debug.Print "Ang1:" + Str(Ang1) + ", Ang2:" + Str(Ang2) + ", Ang3:" + Str(Ang3) + ", Ang4:" + Str(Ang3)
         GoSub NewPoint
         GivePointPlane mX1, mY1, mX2, mY2, Ang1, Ang2, 0, P, , , , Patima
         outP(1).X = P(1).X
         outP(2).X = P(2).X
         outP(1).Y = P(1).Y
         outP(2).Y = P(2).Y
         GoSub NewPoint
         GivePointPlane mX2, mY2, mX3, mY3, Ang2, Ang3, 0, P, , , , Patima
         outP(3).X = P(2).X
         outP(3).Y = P(2).Y
         GoSub NewPoint
         GivePointPlane mX3, mY3, mX4, mY4, Ang3, Ang4, 0, P, , , , Patima
         outP(4).X = P(2).X
         outP(4).Y = P(2).Y
         ''Debug.Print outP(1).X, outP(1).Y, outP(2).X, outP(2).Y, outP(3).X, outP(3).Y, outP(4).X, outP(4).Y
Exit Function
NewPoint:
    mX1 = X1
    mX2 = X2
    mX3 = X3
    mX4 = X4
    mY1 = Y1
    mY2 = Y2
    mY3 = Y3
    mY4 = Y4
Return
End Function


'Õðïëïãéóìüò 4 óçìåßùí óå äåäïìÝíåò óõíôåôáãìÝíåò êáé ãùíßåò
' õðïëïãéæåé ïëá ôá êïììÜôéá ôùí áëïõìéíßùí - ï÷é åðéöáíåéåò
'Function GivePoint(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                   Ang1 As Single, Ang2 As Single, WProfil As Single, POINT() as PointApi, _
                   Optional TPoint As Integer = 1, Optional Patima As Single = 0)
' PointDraw() As PointApi, _

Function GivePoint(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                   Ang1 As Single, Ang2 As Single, WProfil As Single, _
                   PointDraw() As POINTAPI, _
                   Optional WProfil1 As Single = 0, Optional WProfil2 As Single = 0, _
                   Optional Ftero As Single = 0, Optional Patima As Single = 0, _
                   Optional Editing As Single = 0, Optional Scalelable As Single = 0)
         
         Dim P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI, P4 As POINTAPI, P As POINTAPI
         Dim P_1 As POINTAPI, P_2 As POINTAPI, P_3 As POINTAPI, P_4 As POINTAPI
         Dim tp1 As POINTAPI, tp2 As POINTAPI
         Dim TAAS As Single, AngOrig As Single, Distance As Single
         Dim CAng1 As Single, CAng2 As Single
         Dim mX1 As Single, mY1 As Single, mX2 As Single, mY2 As Single
         Dim cX1 As Single, cY1 As Single, cX2 As Single, cY2 As Single
         Dim nX1 As Single, nY1 As Single, nX2 As Single, nY2 As Single
         Dim WInProfil1 As Single, WInProfil2 As Single, wp As Single
         Dim WinProfilA As Single, WInProfilB As Single
         Dim Ft1 As Single, Ft2 As Single, Pt1 As Single, Pt2 As Single
         Dim oldx1 As Single, oldx2 As Single
         oldx1 = X1
         oldx2 = X2
         Y1 = -Y1
         Y2 = -Y2
         X1 = oldx1
         X2 = oldx2
        ' if ang1<>90 then Ft1=
         'ÃÙÍÉÁ ÊÏÐÇÓ 1
         CAng1 = 90 - Ang1
         'ÃÙÍÉÁ ÊÏÐÇÓ 2
         CAng2 = 90 - Ang2
         'ÌÞêïò êïðÞò ðñïößë
         Distance = Dist(X1, Y1, X2, Y2)   'ok ìÞêïò êïðÞò ðñïößë
         
         cX2 = Distance + Editing
         ''Debug.Print "Distance:" + Str(Distance) + " - Ang1:" + Str(cAng1) + " - Ang2:" + Str(cAng2)
         wp = (WProfil + Ftero) '- Patima)
         'ÌÞêïò êáèåôçò ðëåõñÜò ðÜíù óôï ðñïößë
         Ft1 = TheoremAAS(CAng1, Ang1, Ftero)  'ok
         Ft2 = TheoremAAS(CAng2, Ang2, Ftero)  'ok
         Pt1 = TheoremAAS(CAng1, Ang1, Patima)  'ok
         Pt2 = TheoremAAS(CAng2, Ang2, Patima)  'ok
         WInProfil1 = TheoremAAS(CAng1, Ang1, wp)   'ok
         WInProfil2 = TheoremAAS(CAng2, Ang2, wp) '
         WinProfilA = TheoremAAS(Ang1, Ang1, WProfil1) 'ok
         WInProfilB = TheoremAAS(Ang2, Ang2, WProfil2) '
        ' 'Debug.Print "WInProfil1:" + Str(WInProfil1) + " - WInProfil2:" + Str(WInProfil2)
        ' 'Debug.Print "WInProfilA:" + Str(WinProfilA) + " - WInProfilB:" + Str(WInProfilB)
         'Ãùíßá áðü ôï 0,0
         mX1 = X1 + (-X1)
         mY1 = Y1 + (-Y1)
         mX2 = X2 + (-X1)
         mY2 = Y2 + (-Y1)
         AngOrig = GetAngle3P(mX1, mY1, mX2, mY2, mX2, Min(mY1, Abs(mY2))) 'ok
        ' 'Debug.Print "AngOrig:" + Str(AngOrig)
         'Õðïëïãéóìüò óôï 0,0
         cX1 = WinProfilA - Ft1 - Pt1
         cX2 = cX2 - WInProfilB + Ft2 + Pt2
         nX1 = cX1 + WInProfil1
         nY1 = cY1 + -WProfil
         nX2 = cX2 + -WInProfil2 '+ Pt2
         nY2 = cY2 + -WProfil
        ' 'Debug.Print "New Point "
        ' 'Debug.Print Str(cX1) + "," + Str(cY1) + "," + Str(cX2) + "," + Str(cY2)
        ' 'Debug.Print Str(nX1) + "," + Str(nY1) + "," + Str(nX2) + "," + Str(nY2)
         P1.X = cX1 '- Patima
         P1.Y = cY1 + Patima
         P2.X = cX2 '+ Patima
         P2.Y = cY2 + Patima
         P3.X = nX1
         P3.Y = nY1 + Patima
         P4.X = nX2
         P4.Y = nY2 + Patima
        ' 'Debug.Print "Ðáôçìá Point "
        ' 'Debug.Print Str(P1.X) + "," + Str(P1.Y) + "," + Str(P2.X) + "," + Str(P2.Y)
        ' 'Debug.Print Str(P3.X) + "," + Str(P3.Y) + "," + Str(P4.X) + "," + Str(P4.Y)
'         If Patima <> 0 Then
'             tp1 = ExtendPoint(P2.X, P2.Y, P1.X, P1.Y, nX1, nY1, cX1, cY1)
'             tp2 = ExtendPoint(P1.X, P1.Y, P2.X, P2.Y, nX2, nY2, cX2, cY2)
'             P1 = tp1
'             P2 = tp2
'             'Debug.Print "Ìå ÐÜôçìá"
'             'Debug.Print Str(P1.X) + "," + Str(P1.Y) + "," + Str(P2.X) + "," + Str(P2.Y)
'             'Debug.Print Str(P3.X) + "," + Str(P3.Y) + "," + Str(P4.X) + "," + Str(P4.Y)
'         End If
         'ÐåñéóôñïöÞ ôïõ ôåìá÷ßïõ óôï 0,0
         P_1 = Rotate(P1.X, P1.Y, AngOrig)  'ÏÊ
         P_2 = Rotate(P2.X, P2.Y, AngOrig)  'ÏÊ
         P_3 = Rotate(P3.X, P3.Y, AngOrig)  'ÏÊ
         P_4 = Rotate(P4.X, P4.Y, AngOrig)  'ÏÊ
        ' 'Debug.Print "ROTATE Point "
        ' 'Debug.Print Str(P_1.X) + "," + Str(P_1.Y) + "," + Str(P_2.X) + "," + Str(P_2.Y)
        ' 'Debug.Print Str(P_3.X) + "," + Str(P_3.Y) + "," + Str(P_4.X) + "," + Str(P_4.Y)
         
'         'Ìåôáêßíçóç ôùí óçìåßùí óôá ÔåëéêÜ Óçìåßá
         P1.X = P_1.X + X1
         P1.Y = -(P_1.Y + Y1)
         P2.X = P_2.X + X1
         P2.Y = -(P_2.Y + Y1)
         P3.X = P_4.X + X1
         P3.Y = -(P_4.Y + Y1)
         P4.X = P_3.X + X1
         P4.Y = -(P_3.Y + Y1)
         Y1 = -Y1
         Y2 = -Y2
         ReDim PointDraw(1 To 4) As POINTAPI
         PointDraw(1).X = P1.X
         PointDraw(2).X = P2.X
         PointDraw(3).X = P3.X
         PointDraw(4).X = P4.X
         PointDraw(1).Y = P1.Y
         PointDraw(2).Y = P2.Y
         PointDraw(3).Y = P3.Y
         PointDraw(4).Y = P4.Y
         ''Debug.Print "MOVE point:" + Str(p1.X) + "," + Str(p1.Y) + "," + Str(p2.X) + "," + Str(p2.Y) + "," + Str(p3.X) + "," + Str(p3.Y) + "," + Str(P4.X) + "," + Str(P4.Y)
         ''Debug.Print "Discance:" + Str(Dist(p1.X, p1.Y, p2.X, p2.Y))
End Function
Private Function GivePointPlane(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                        Ang1 As Single, Ang2 As Single, WProfil As Single, _
                        PointDraw() As POINTAPI, _
                        Optional WProfil1 As Single = 0, Optional WProfil2 As Single = 0, _
                        Optional TPoint As Integer = 1, Optional Patima As Single = 0, _
                        Optional Editing As Single = 0, Optional Scalelable As Single = 0)
         
         Dim P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI, P4 As POINTAPI, P As POINTAPI
         Dim P_1 As POINTAPI, P_2 As POINTAPI, P_3 As POINTAPI, P_4 As POINTAPI
         Dim tp1 As POINTAPI, tp2 As POINTAPI
         Dim TAAS As Single, AngOrig As Single, Distance As Single
         Dim CAng1 As Single, CAng2 As Single
         Dim mX1 As Single, mY1 As Single, mX2 As Single, mY2 As Single
         Dim cX1 As Single, cY1 As Single, cX2 As Single, cY2 As Single
         Dim nX1 As Single, nY1 As Single, nX2 As Single, nY2 As Single
         Dim WInProfil1 As Single, WInProfil2 As Single, wp As Single
         Dim WinProfilA As Single, WInProfilB As Single
         Y1 = -Y1
         Y2 = -Y2
         'ÃÙÍÉÁ ÊÏÐÇÓ 1
         CAng1 = 90 - Ang1
         'ÃÙÍÉÁ ÊÏÐÇÓ 2
         CAng2 = 90 - Ang2
         'ÌÞêïò êïðÞò ðñïößë
         Distance = Dist(X1, Y1, X2, Y2)   'ok ìÞêïò êïðÞò ðñïößë
         cX2 = Distance + Editing
'         'Debug.Print "Distance:" + Str(Distance) + " - Ang1:" + Str(cAng1) + " - Ang2:" + Str(cAng2)
         wp = (WProfil - Patima)
         'ÌÞêïò êáèåôçò ðëåõñÜò ðÜíù óôï ðñïößë
         WInProfil1 = TheoremAAS(CAng1, Ang1, wp) 'ok
         WInProfil2 = TheoremAAS(CAng2, Ang2, wp) '
'         'Debug.Print "WInProfil1:" + Str(WInProfil1) + " - WInProfil2:" + Str(WInProfil2)
         'Ãùíßá áðü ôï 0,0
         mX1 = X1 + (-X1)
         mY1 = Y1 + (-Y1)
         mX2 = X2 + (-X1)
         mY2 = Y2 + (-Y1)
         AngOrig = GetAngle3P(mX1, mY1, mX2, mY2, Max(mX1, mX2), Min(mY1, Abs(mY2))) 'ok
'         'Debug.Print "AngOrig:" + Str(AngOrig)
         'Õðïëïãéóìüò óôï 0,0
         nX1 = cX1 - WInProfil1
         nY1 = cY1 - -Patima
         nX2 = cX2 + WInProfil2
         nY2 = cY2 - -Patima
'         'Debug.Print "New Point "
'         'Debug.Print Str(nX1) + "," + Str(nY1) + "," + Str(nX2) + "," + Str(nY2)
         P1.X = nX1
         P1.Y = nY1
         P2.X = nX2
         P2.Y = nY2
'         'Debug.Print "Ðáôçìá Point "
'         'Debug.Print Str(P1.X) + "," + Str(P1.Y) + "," + Str(P2.X) + "," + Str(P2.Y)
         'ÐåñéóôñïöÞ ôïõ ôåìá÷ßïõ óôï 0,0
         P_1 = Rotate(P1.X, P1.Y, AngOrig) 'ÏÊ
         P_2 = Rotate(P2.X, P2.Y, AngOrig) 'ÏÊ
         P1 = P_1
         P2 = P_2
'         'Debug.Print "ROTATE Point "
'         'Debug.Print Str(P_1.X) + "," + Str(P_1.Y) + "," + Str(P_2.X) + "," + Str(P_2.Y)
         
        'Debug.Print
'         'Ìåôáêßíçóç ôùí óçìåßùí óôá ÔåëéêÜ Óçìåßá
         P1.X = P1.X + X1
         P1.Y = -(P1.Y + Y1)
         P2.X = P2.X + X1
         P2.Y = -(P2.Y + Y1)
         
         ReDim PointDraw(1 To 2)
         PointDraw(1).X = P1.X
         PointDraw(2).X = P2.X
         PointDraw(1).Y = P1.Y
         PointDraw(2).Y = P2.Y
       '  'Debug.Print "MOVE point:" + Str(P1.X) + "," + Str(P1.Y) + "," + Str(P2.X) + "," + Str(P2.Y)
        ' 'Debug.Print "Discance:" + Str(Dist(P1.X, P1.Y, P2.X, P2.Y))
End Function

 Function GivePointVMulti(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                            Ang1 As Single, WProfil As Single, _
                            Bima As Single, IdPoint As Integer, Editing As Single, PointDraw() As POINTAPI)
         
         Dim P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI, P4 As POINTAPI, P As POINTAPI
         Dim P_1 As POINTAPI, P_2 As POINTAPI, P_3 As POINTAPI, P_4 As POINTAPI
         Dim tp1 As Single, tp2 As Single
         Dim AngOrig As Single, Distance As Single
         Dim CAng1 As Single, CAng2 As Single
         Dim mX1 As Single, mY1 As Single, mX2 As Single, mY2 As Single
         Dim cX1 As Single, cY1 As Single, cX2 As Single, cY2 As Single
         Dim nX1 As Single, nY1 As Single, nX2 As Single, nY2 As Single
         Dim wp As Single
         
         Y1 = -Y1
         Y2 = -Y2
         'ÃÙÍÉÁ ÊÏÐÇÓ 1
         CAng1 = 90 - Ang1
         'ÃÙÍÉÁ ÊÏÐÇÓ 2
         'ÌÞêïò êïðÞò ðñïößë
         Distance = Dist(X1, Y1, X2, Y2)   'ok ìÞêïò êïðÞò ðñïößë
         cX1 = (Bima * IdPoint)
         'cX2 = (Distance + Editing) - (Bima * IdPoint)
'    'Debug.Print "Distance:" + Str(Distance) + " - Ang1:" + Str(cAng1) '+ " - Ang2:" + Str(cAng2)
         wp = (WProfil) ' - Patima)
         'ÌÞêïò êáèåôçò ðëåõñÜò ðÜíù óôï ðñïößë
         If Ang1 <> 0 Then
             WInProfil1 = TheoremAAS(CAng1, Ang1, wp) 'ok
'            'Debug.Print "WInProfil1:" + Str(WInProfil1)
         End If
      
       'Ãùíßá áðü ôï 0,0
       mX1 = X1 + (-X1)
       mY1 = Y1 + (-Y1)
       mX2 = X2 + (-X1)
       mY2 = Y2 + (-Y1)
       AngOrig = GetAngle3P(mX1, mY1, mX2, mY2, Max(mX1, mX2), Min(mY1, Abs(mY2))) 'ok
'       'Debug.Print "AngOrig:" + Str(AngOrig)
       'Õðïëïãéóìüò óôï 0,0
       cX1 = (Bima) - (wp / 2)  '- WInProfil1
       cX2 = (Bima) + (wp / 2)  '+ WInProfil1
       
'        'Debug.Print "New Point "
'        'Debug.Print Str(cX1) + "," + Str(cY1) + "," + Str(cX2) + "," + Str(cY2)
        
       P1.X = cX1
       P1.Y = cY1
       P2.X = cX2
       P2.Y = cY2
       ''Debug.Print "ÂÇÌÁ", cX1, cY1, cX2, cY2
       'Stop
       'ÐåñéóôñïöÞ ôïõ ôåìá÷ßïõ óôï 0,0
       P_1 = Rotate(P1.X, P2.Y, AngOrig) 'ÏÊ
       P_2 = Rotate(P2.X, P2.Y, AngOrig) 'ÏÊ
'      'Debug.Print "ROTATE Point "
'      'Debug.Print Str(P_1.X) + "," + Str(P_1.Y) + "," + Str(P_2.X) + "," + Str(P_2.Y)
'       'Ìåôáêßíçóç ôùí óçìåßùí óôá ÔåëéêÜ Óçìåßá
       P1.X = P_1.X + X1
       P1.Y = -(P_1.Y + Y1)
       P2.X = P_2.X + X1
       P2.Y = -(P_2.Y + Y1)
        
       Y1 = -Y1
       Y2 = -Y2
        
       ReDim PointDraw(1 To 2)
       PointDraw(1).X = P1.X
       PointDraw(2).X = P2.X
       PointDraw(1).Y = P1.Y
       PointDraw(2).Y = P2.Y
       
'       'Debug.Print "MOVE point:" + Str(P1.X) + "," + Str(P1.Y) + "," + Str(P2.X) + "," + Str(P2.Y) '+ "," + Str(P3.X) + "," + Str(P3.Y) + "," + Str(P4.X) + "," + Str(P4.Y)
'       'Debug.Print "Discance:" + Str(Dist(P1.X, P1.Y, P2.X, P2.Y))
End Function

'õðïëïãéóìïò ìÞêïõò ôáö
'åëåã÷ï áí èÝëåé áëëáãÞ óôá Õ
Function GiveVerticalLenght(TypePoint As Integer, cX1 As Single, cY1 As Single, _
                            X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                            X3 As Single, Y3 As Single, X4 As Single, Y4 As Single, _
                            Ang1 As Single, Ang2 As Single, _
                            WProfil As Single, P() As POINTAPI, _
                            Optional Editing As Single = 0, _
                            Optional Scalelable As Single = 0) As Single
         
         Dim nAng1 As Single, nAng2 As Single ', P() as PointApi
         Dim d1 As Single, D2 As Single
        
         Select Case TypePoint
         Case 1 'vertical
               GiveVertical TypePoint, cX1, cY1, X1, Y1, X2, Y2, X3, Y3, X4, Y4, WProfil, P
               d1 = Dist(P(1).X, P(1).Y, P(2).X, P(2).Y)
               D2 = Dist(P(3).X, P(3).Y, P(4).X, P(4).Y)
               If d1 > D2 Then GiveVerticalLenght = d1 Else GiveVerticalLenght = D2
               nAng1 = GetInternalAngle3P(P(1).X, P(1).Y, P(2).X, P(2).Y, P(4).X, P(4).Y)
               nAng2 = GetInternalAngle3P(P(2).X, P(2).Y, P(3).X, P(3).Y, P(1).X, P(1).Y)
         Case 2
               GiveVertical TypePoint, cX1, cY1, X1, Y1, X2, Y2, X3, Y3, X4, Y4, WProfil, P
               d1 = Dist(P(1).X, P(1).Y, P(4).X, P(4).Y)
               D2 = Dist(P(2).X, P(2).Y, P(3).X, P(3).Y)
               If d1 > D2 Then GiveVerticalLenght = d1 Else GiveVerticalLenght = D2
               nAng1 = GetInternalAngle3P(P(1).X, P(1).Y, P(2).X, P(2).Y, P(4).X, P(4).Y)
               nAng2 = GetInternalAngle3P(P(2).X, P(2).Y, P(3).X, P(3).Y, P(1).X, P(1).Y)
         End Select
         GiveVerticalLenght = GiveVerticalLenght + Editing + Scalelable
        ''Debug.Print "Ang1:", nAng1, "Ang2:", nAng2
         Ang1 = nAng1
         Ang2 = nAng2
         
End Function

'õðïëïãéóìïò óçìåßùí ôáö ãéá ó÷åäßáóç
'åëåã÷ï áí èÝëåé áëëáãÞ óôá Õ
Function GiveVertical(TypePoint As Integer, cX1 As Single, cY1 As Single, _
                      X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                      X3 As Single, Y3 As Single, X4 As Single, Y4 As Single, _
                      WProfil As Single, NewPoint() As POINTAPI)
                      
         Dim mp1 As POINTAPI, mp2 As POINTAPI, mp3 As POINTAPI, mp4 As POINTAPI, P() As POINTAPI
         Dim mX1 As Single, mY1 As Single, mX2 As Single, mY2 As Single
         Dim DistPoint1 As Single, DistPoint2 As Single, DistX As Single, DistY As Single
         Dim Dist1 As Single, Dist2 As Single, DistLine As Single
         Dim AngOrig  As Single, Bima As Single
         Dim nAng1 As Single, nAng2 As Single
         ReDim NewPoint(1 To 4) As POINTAPI
         
         Select Case TypePoint
         Case 1 'vertical
             mp1 = ExtendPoint(X1, Y1, X4, Y4, cX1 - (WProfil / 2), Max(Y1, Y2), cX1 - (WProfil / 2), Min(Y1, Y2) - 10)
             mp4 = ExtendPoint(X1, Y1, X4, Y4, cX1 + (WProfil / 2), Max(Y1, Y2), cX1 + (WProfil / 2), Min(Y1, Y2) - 10)
           '  GivePointV X1, Y1, X4, Y4, nAng1, WProfil, cX1, Editing, p
           ' 'Debug.Print "GiveVertical:", p(1).X, p(1).Y, p(2).X, p(2).Y
'            NEWPOINT(1).X = p(1).X
'            NEWPOINT(1).Y = p(1).Y
'            NEWPOINT(2).X = p(2).X
'            NEWPOINT(2).Y = p(2).Y
             NewPoint(1).X = mp1.X
             NewPoint(1).Y = mp1.Y
             NewPoint(4).X = mp4.X
             NewPoint(4).Y = mp4.Y
             
             mp2 = ExtendPoint(X2, Y2, X3, Y3, cX1 - (WProfil / 2), Max(Y2, Y4), cX1 - (WProfil / 2), Min(Y2, Y3))
             mp3 = ExtendPoint(X2, Y2, X3, Y3, cX1 + (WProfil / 2), Max(Y2, Y4), cX1 + (WProfil / 2), Min(Y2, Y3))
             'GivePointV X2, Y2, X3, Y3, nAng2, WProfil, cX1, Editing, p
'            'Debug.Print "GiveVertical:", p(1).X, p(1).Y, p(2).X, p(2).Y
'            NEWPOINT(3).X = p(1).X
'            NEWPOINT(3).Y = p(1).Y
'            NEWPOINT(4).X = p(2).X
'            NEWPOINT(4).Y = p(2).Y
             NewPoint(2).X = mp2.X
             NewPoint(2).Y = mp2.Y
             NewPoint(3).X = mp3.X
             NewPoint(3).Y = mp3.Y
             'Stop
             ''Debug.Print Dist(mp1.X, mp1.Y, mp2.X, mp2.Y)
             ''Debug.Print Dist(mp3.X, mp3.Y, mp4.X, mp4.Y)
         Case 2 'orizontal
'             nAng1 = GetAngle3P(X3, Y3, X4, Y4, Max(X3, X4), Min(Y3, Abs(Y4))) - 180
'             nAng2 = GetAngle3P(X2, Y2, X1, Y1, Max(X1, X2), Min(Y1, Abs(Y2))) - 180
'             GivePointV X4, Y4, X3, Y3, nAng1, WProfil, cY1, Editing, p
'             'Debug.Print "GiveVertical:", p(1).X, p(1).Y, p(2).X, p(2).Y
'             NEWPOINT(1).X = p(1).X
'             NEWPOINT(1).Y = p(1).Y
'             NEWPOINT(2).X = p(2).X
'             NEWPOINT(2).Y = p(2).Y
'             GivePointV X1, Y1, X2, Y2, nAng2, WProfil, cY1, Editing, p
'             'Debug.Print "GiveVertical:", p(1).X, p(1).Y, p(2).X, p(2).Y
'             NEWPOINT(3).X = p(1).X
'             NEWPOINT(3).Y = p(1).Y
'             NEWPOINT(4).X = p(2).X
'             NEWPOINT(4).Y = p(2).Y
             mp1 = ExtendPoint(X3, Y3, X4, Y4, Max(Y3, Y4), cY1 - (WProfil / 2), Min(Y3, Y4) - 10, cY1 - (WProfil / 2))
             mp4 = ExtendPoint(X3, Y3, X4, Y4, Max(Y3, Y4), cY1 + (WProfil / 2), Min(Y3, Y4) - 10, cY1 + (WProfil / 2))
             mp2 = ExtendPoint(X1, Y1, X2, Y2, Max(Y1, Y2), cY1 - (WProfil / 2), Min(Y1, Y2) - 10, cY1 - (WProfil / 2))
             mp3 = ExtendPoint(X1, Y1, X2, Y2, Max(Y1, Y2), cY1 + (WProfil / 2), Min(Y1, Y2) - 10, cY1 + (WProfil / 2))
             NewPoint(3).X = mp1.X
             NewPoint(3).Y = mp1.Y
             NewPoint(2).X = mp2.X
             NewPoint(2).Y = mp2.Y
             NewPoint(1).X = mp3.X
             NewPoint(1).Y = mp3.Y
             NewPoint(4).X = mp4.X
             NewPoint(4).Y = mp4.Y
'             'Debug.Print "-", mp1.X, mp1.Y, mp2.X, mp2.Y
'             'Debug.Print "-", mp3.X, mp3.Y, mp4.X, mp4.Y
'             'Debug.Print Dist(mp1.X, mp1.Y, mp2.X, mp2.Y)
'             'Debug.Print Dist(mp3.X, mp3.Y, mp4.X, mp4.Y)
         End Select
                   
End Function

'Private Function GivePointV(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
'                            Ang1 As Single, WProfil As Single, _
'                            nX1 As Single, Editing As Single, PointDraw() as PointApi)
'
'         Dim P1 as PointApi, P2 as PointApi, P3 as PointApi, P4 as PointApi, p as PointApi
'         Dim P_1 as PointApi, P_2 as PointApi, P_3 as PointApi, P_4 as PointApi
'         Dim tp1 as PointApi, tp2 as PointApi
'         Dim AngOrig As Single, Distance As Single
'         Dim cAng1 As Single, cAng2 As Single
'         Dim mX1 As Single, mY1 As Single, mX2 As Single, mY2 As Single
'         Dim cX1 As Single, cY1 As Single, cX2 As Single, cY2 As Single
'         'Dim nX1 As Single, nY1 As Single, nX2 As Single, nY2 As Single
'         Dim wp As Single
'
'         Y1 = -Y1
'         Y2 = -Y2
'         'ÃÙÍÉÁ ÊÏÐÇÓ 1
'         cAng1 = 90 - Ang1
'         'ÃÙÍÉÁ ÊÏÐÇÓ 2
'         'ÌÞêïò êïðÞò ðñïößë
'         Distance = Dist(X1, Y1, X2, Y2)   'ok ìÞêïò êïðÞò ðñïößë
'        'cX1 = (Bima * IdPoint)
'        cX2 = Editing
'        'Debug.Print "Distance:" + Str(Distance) + " - Ang1:" + Str(cAng1) '+ " - Ang2:" + Str(cAng2)
'         wp = (WProfil) ' - Patima)
'         'ÌÞêïò êáèåôçò ðëåõñÜò ðÜíù óôï ðñïößë
'       ' cX1 =
'         If Ang1 <> 0 Then
'             WInProfil1 = TheoremAAS2(cAng1, Ang1, wp)  'ok
'             'Debug.Print "WInProfil1:" + Str(WInProfil1)
'         End If
'
'       'Ãùíßá áðü ôï 0,0
'       mX1 = X1 + (-X1)
'       mY1 = Y1 + (-Y1)
'       mX2 = X2 + (-X1)
'       mY2 = Y2 + (-Y1)
'       AngOrig = GetAngle3P(mX1, mY1, mX2, mY2, Max(mX1, mX2), Min(mY1, Abs(mY2))) 'ok
'       'Debug.Print "AngOrig:" + Str(AngOrig)
'         'Õðïëïãéóìüò óôï 0,0
'      cX1 = cX1 - (WInProfil1 / 2)
'      cX2 = cX2 + (WInProfil1 / 2)
'        'Debug.Print "New Point "
'        'Debug.Print Str(cX1) + "," + Str(cY1) + "," + Str(cX2) + "," + Str(cY2)
'
'       P1.X = cX1
'       P1.Y = cY1
'       P2.X = cX2
'       P2.Y = cY2
'
'       'ÐåñéóôñïöÞ ôïõ ôåìá÷ßïõ óôï 0,0
'       P_1 = Rotate(P1, AngOrig) 'ÏÊ
'       P_2 = Rotate(P2, AngOrig) 'ÏÊ
'      'Debug.Print "ROTATE Point "
'      'Debug.Print Str(P_1.X) + "," + Str(P_1.Y) + "," + Str(P_2.X) + "," + Str(P_2.Y)
''       'Ìåôáêßíçóç ôùí óçìåßùí óôá ÔåëéêÜ Óçìåßá
'       P1.X = P_1.X + X1
'       P1.Y = P_1.Y + Y1
'       P2.X = P_2.X + X1
'       P2.Y = P_2.Y + Y1
'
'       Y1 = -Y1
'       Y2 = -Y2
'
'       ReDim PointDraw(1 To 2)
'       PointDraw(1).X = P1.X
'       PointDraw(2).X = P2.X
'       PointDraw(1).Y = P1.Y
'       PointDraw(2).Y = P2.Y
'
'       'Debug.Print "MOVE point:" + Str(P1.X) + "," + Str(P1.Y) + "," + Str(P2.X) + "," + Str(P2.Y) '+ "," + Str(P3.X) + "," + Str(P3.Y) + "," + Str(P4.X) + "," + Str(P4.Y)
'       'Debug.Print "Discance:" + Str(Dist(P1.X, P1.Y, P2.X, P2.Y))
'End Function


'PointDraw() as PointApi, _
'
Function GiveVerticalMulti(TypePoint As Integer, MaxPoint As Integer, IdPoint As Integer, _
                      X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                      X3 As Single, Y3 As Single, X4 As Single, Y4 As Single, _
                      Ang1 As Single, Ang2 As Single, WProfil As Single, _
                      NewPoint() As POINTAPI, _
                      Optional Editing As Single = 0, Optional Scalelable As Single = 0)
                      
         Dim mp1 As POINTAPI, mp2 As POINTAPI, P() As POINTAPI
         Dim cX1 As Single, cY1 As Single, cX2 As Single, cY2 As Single, cX3 As Single, cY3 As Single, cX4 As Single, cY4 As Single
         Dim mX1 As Single, mY1 As Single, mX2 As Single, mY2 As Single
         Dim DistPoint1 As Single, DistPoint2 As Single, DistX As Single, DistY As Single
         Dim Dist1 As Single, Dist2 As Single, DistLine As Single
         Dim AngOrig  As Single, Bima As Single
         Dim nAng1 As Single, nAng2 As Single
         Dim tp1 As Single, tp2 As Single
         ReDim NewPoint(1 To 4) As POINTAPI
         
        ' If MaxPoint = IdPoint Then tp2
         
         Select Case TypePoint
         Case 1 'vertical
             'Bima = Max(Dist(X3, Y3, X4, Y4), Dist(X1, Y1, X2, Y2)) / (MaxPoint + 1) + (MaxPoint - 1) * WProfil + WProfil / 2
             Bima = ((Max(Dist(X1, Y1, X4, Y4), Dist(X2, Y2, X3, Y3)) - (WProfil * MaxPoint)) / (MaxPoint + 1))
             Bima = Bima * IdPoint
             Bima = Bima + ((IdPoint - 1) * WProfil) + (WProfil / 2)
             nAng1 = GetAngle3P(X1, Y1, X4, Y4, Max(X1, X2), Min(Y3, Abs(Y4)))
             nAng2 = GetAngle3P(X2, Y2, X3, Y3, Max(X2, X3), Min(Y2, Abs(Y3)))
             GivePointVMulti X1, Y1, X4, Y4, nAng1, WProfil, Bima, IdPoint, Editing, P
            ' 'Debug.Print "GiveVertical:", p(1).X, p(1).Y, p(2).X, p(2).Y
             NewPoint(1).X = P(1).X
             NewPoint(1).Y = P(1).Y
             NewPoint(4).X = P(2).X
             NewPoint(4).Y = P(2).Y
             GivePointVMulti X2, Y2, X3, Y3, nAng2, WProfil, Bima, IdPoint, Editing, P
            ' 'Debug.Print "GiveVertical:", p(1).X, p(1).Y, p(2).X, p(2).Y
             NewPoint(2).X = P(1).X
             NewPoint(2).Y = P(1).Y
             NewPoint(3).X = P(2).X
             NewPoint(3).Y = P(2).Y
            ' Stop
         Case 2 'orizontal
             nAng1 = GetAngle3P(X3, Y3, X4, Y4, Max(X3, X4), Min(Y3, Abs(Y4))) - 180
             nAng2 = GetAngle3P(X2, Y2, X1, Y1, Max(X1, X2), Min(Y1, Abs(Y2))) - 180
             'Bima = Max(Dist(X3, Y3, X4, Y4), Dist(X1, Y1, X2, Y2)) / (MaxPoint + 1) + (MaxPoint - 1) * WProfil + WProfil / 2
             Bima = ((Max(Dist(X3, Y3, X4, Y4), Dist(X1, Y1, X2, Y2)) - (WProfil * MaxPoint)) / (MaxPoint + 1))
             Bima = Bima * IdPoint
             Bima = Bima + ((IdPoint - 1) * WProfil) + (WProfil / 2)
             GivePointVMulti X4, Y4, X3, Y3, nAng1, WProfil, Bima, IdPoint, Editing, P
            ' 'Debug.Print "GiveVertical:", P(1).X, P(1).Y, P(2).X, P(2).Y
             NewPoint(4).X = P(1).X
             NewPoint(4).Y = P(1).Y
             NewPoint(3).X = P(2).X
             NewPoint(3).Y = P(2).Y
             GivePointVMulti X1, Y1, X2, Y2, nAng2, WProfil, Bima, IdPoint, Editing, P
            ' 'Debug.Print "GiveVertical:", P(1).X, P(1).Y, P(2).X, P(2).Y
             NewPoint(1).X = P(1).X
             NewPoint(1).Y = P(1).Y
             NewPoint(2).X = P(2).X
             NewPoint(2).Y = P(2).Y
         End Select
       ' 'Debug.Print "GiveVertical:", NEWPOINT(1).X, NEWPOINT(1).Y, NEWPOINT(2).X, NEWPOINT(2).Y, NEWPOINT(3).X, NEWPOINT(3).Y, NEWPOINT(4).X, NEWPOINT(4).Y

End Function

Function ExtendPoint(nX1 As Single, nY1 As Single, nX2 As Single, nY2 As Single, _
                     mX1 As Single, mY1 As Single, mX2 As Single, mY2 As Single) As POINTAPI
                     
        Dim m1 As Single, m2 As Single, b1 As Single, b2 As Single
        Dim X1 As Single, Y As Single, IsVertical As Boolean
'       'Debug.Print
'        'Debug.Print "Extend Point"
'        'Debug.Print Str(nX1) + "," + Str(nY1) + "," + Str(nX2) + "," + Str(nY2)
'        'Debug.Print Str(mX1) + "," + Str(mY1) + "," + Str(mX2) + "," + Str(mY2)
        'm = Y1 - Y2 / X1 - X2
        'b= y - m x
        'x= b2-b1/m1-m2
        'y=mx+b
        m1 = 0
        m2 = 0
        b1 = 0
        b2 = 0
        X = 0
        Y = 0
'        m1 = SlopeLine(nX1, nY1, nX2, nY2)
'        m2 = SlopeLine(mX1, mY1, mX2, mY2)
         
        If nX1 = nX2 Then 'êáèåôï
           X = nX1
           m1 = 0
           IsVertical = True
        Else
           If (nX1 - nX2) <> 0 Then m1 = (nY1 - nY2) / (nX1 - nX2) Else m1 = 0
        End If
        If mX1 = mX2 Then
           X = mX1
           m2 = 0
           IsVertical = True
        Else
           If (mX1 - mX2) <> 0 Then m2 = (mY1 - mY2) / (mX1 - mX2) Else m2 = 0
        End If
        
        If (m1 <> 0 And m2 <> 0) And (m1 = m2) Then MsgBox "ÐáñÜëëçëåò", vbCritical: Exit Function
        
        b1 = nY1 - (m1 * nX1)
        b2 = mY2 - (m2 * nX2)
'        b1 = InterceptLine(nX1, nY1, nX2, nY2)
'        b2 = InterceptLine(mX1, mY1, mX2, mY2)
        
        If (m1 - m2) <> 0 And IsVertical = False Then X = (b2 - b1) / (m1 - m2) 'Else X = 0
        Y = m1 * X + b1
        'Y = m2 * X + b2
        
        If nY1 = nY2 Then Y = nY1
        
        If mY1 = mY2 Then Y = mY1
        
        'Y = m2 * X + b2
'        Debug.Print
'        'Debug.Print Str(X) + "=[" + Trim(Str(b1)) + "-" + Trim(Str(b2)) + "]/[" + Trim(Str(m1)) + "-" + Trim(Str(m2)) + "]"
'        'Debug.Print Str(Y) + "=" + Str(m1) + " * " + Trim(Str(X)) + " + " + Trim(Str(b1))
'
        ExtendPoint.X = Format(X, "0.0")
        ExtendPoint.Y = Format(Y, "0.0")
        
End Function

'2 ãùíßåò ãíùóôÝò - ìéá ðëåõñÜ ãíùóôÞ óôéò 90 ìïßñåò
'TheoremAAS(54.56,35.54,50)=70
'         /|
'        / |
'       /36|
'      /   | 70
'     /    |
'    /54  _|
'   /____|_|
'     50
Function TheoremAAS(Ang1 As Single, Ang2 As Single, WidthProfil As Single) As Single
         If Ang2 <> 0 Then
         TheoremAAS = WidthProfil * (Sin(Rad(Ang1)) / Sin(Rad(Ang2))) ', "0.00")
         'TheoremAAS = WP * Sin(TheoremAAS) * (Cot(Rad(Ang1)) + Cot(Rad(Ang2)))
         Else
         TheoremAAS = 0
         End If
End Function

Function TheoremAAS2(Ang1 As Single, Ang2 As Single, WidthProfil As Single) As Single
         TheoremAAS2 = WidthProfil * Sin(Rad(Ang1)) * (Cot(Rad(Ang2)) + Cot(Rad(Ang1))) / 10
         'TheoremAAS = WP * Sin(TheoremAAS) * (Cot(Rad(Ang1)) + Cot(Rad(Ang2)))
End Function

Function Cot(xRad As Single) As Single
      Cot = 1 / Tan(xRad)
End Function

Public Function GetAngle(ByVal X1 As Single, ByVal Y1 As Single, X2 As Single, Y2 As Single) As Single
'
Dim G As Single, X As Single, Y As Single ', Pi As Double
'Pi = 3.14159265358979
'
X = X1 - X2
Y = Y1 - Y2
Select Case X
Case Is > 0
    If Y >= 0 Then
        G = Atn(Y / X)
    Else
        G = Atn(Y / X) + 2 * PI
    End If
Case 0
    If Y >= 0 Then
        G = PI / 2
    Else
        G = 3 * PI / 2
    End If
Case Is <= 0
    G = Atn(Y / X) + PI
End Select

G = Degrees(G) ' (g * 180 / Pi)
G = G + 270
If G > 360 Then
    G = G - 360
End If
GetAngle = G
'
End Function

'Function MidPoint(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As POINTAPI
'        Dim cm As POINTAPI
'        cm.X = (X1 + X2) / 2
'        cm.Y = (Y1 + Y2) / 2
'        MidPoint = cm
'End Function
'êëéóç ãñáììÞò m
'Function SlopeLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
'         If (X1 - X2) <> 0 Then
'             SlopeLine = (Y1 - Y2) / (X1 - X2)
'         End If
'End Function
''ôïìÞ ãñáììÞò b
'Function InterceptLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single)
'         InterceptLine = Y1 - SlopeLine(X1, Y1, X2, Y2) * X1
'End Function
'y = mx + â.
Function MakeLine()
      
End Function

'Function PointInLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, Lg As Single) As PointApi
'         Dim P As PointApi, Ang1 As Single, Ang2 As Single
'         Dim A As Single, B As Single, C As Single
'         Y1 = -Y1
'         Y2 = -Y2
'         If X1 = X2 And Y1 <> Y2 Then
'            P.X = X1
'            P.Y = Y2 - Y1
'         ElseIf X1 <> X2 And Y1 = Y2 Then
'            P.X = X2 - X1
'            P.Y = Y1
'         ElseIf X1 = X2 And Y1 = Y2 Then
'            P.X = X1
'            P.Y = Y1
'         ElseIf X1 <> X2 And Y1 <> Y2 Then
'            Ang1 = GetInternalAngle3P(X1, Y1, X1, Y2, X2, Y2)
'            Ang2 = 180 - Ang1 - 90
'            A = Rad(Ang1)
'            B = Rad(Ang2)
'            C = Rad(180 - Ang1 - Ang2)
'            P.Y = -(Sin(B) / Sin(C) * Lg)
'            P.X = -(Sin(A) / Sin(C) * Lg)
'         End If
'         Y1 = -Y1
'         Y2 = -Y2
'         PointInLine = P
'End Function
'
Public Function PointInLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, lg As Single) As POINTAPI
         
         Dim P1 As POINTAPI, P2 As POINTAPI, P As POINTAPI
         Dim P_1 As POINTAPI, P_2 As POINTAPI, P_3 As POINTAPI, P_4 As POINTAPI
         Dim tp1 As POINTAPI, tp2 As POINTAPI
         Dim TAAS As Single, AngOrig As Single, Distance As Single
         Dim CAng1 As Single, CAng2 As Single
         Dim mX1 As Single, mY1 As Single, mX2 As Single, mY2 As Single
         Dim cX1 As Single, cY1 As Single, cX2 As Single, cY2 As Single
         Dim nX1 As Single, nY1 As Single, nX2 As Single, nY2 As Single
         Dim WInProfil1 As Single, WInProfil2 As Single, wp As Single
         Dim WinProfilA As Single, WInProfilB As Single
         Dim Ft1 As Single, Ft2 As Single
         Dim oldx1 As Single, oldx2 As Single
         oldx1 = X1
         oldx2 = X2
         Y1 = -Y1
         Y2 = -Y2
         X1 = oldx1
         X2 = oldx2
        
         'ÃÙÍÉÁ ÊÏÐÇÓ 1
        ' cAng1 = 90 - Ang1
         'ÃÙÍÉÁ ÊÏÐÇÓ 2
        ' cAng2 = 90 - Ang2
         'ÌÞêïò êïðÞò ðñïößë
         Distance = Dist(X1, Y1, X2, Y2)   'ok ìÞêïò êïðÞò ðñïößë
'         'Debug.Print "Distance:" + Str(Distance)
         cX2 = lg '
         
         CAng1 = GetInternalAngle3P(X1, Y1, X1, Y2, X2, Y2)
         CAng2 = 180 - Ang1 - 90
 
         'ÌÞêïò êáèåôçò ðëåõñÜò ðÜíù óôï ðñïößë
         Wv = TheoremAAS(CAng1, CAng2, lg)   'ok
'        'Debug.Print "WInProfil1:" + Str(Wv)

         'Ãùíßá áðü ôï 0,0
         mX1 = X1 + (-X1)
         mY1 = Y1 + (-Y1)
         mX2 = X2 + (-X1)
         mY2 = Y2 + (-Y1)
         AngOrig = GetAngle3P(mX1, mY1, mX2, mY2, Max(mX1, mX2), Min(mY1, Abs(mY2))) 'ok
'         'Debug.Print "AngOrig:" + Str(AngOrig)
         'Õðïëïãéóìüò óôï 0,0
         cX1 = 0 'WinProfilA - Ft1
         cX2 = cX2 '- WInProfilB + Ft2
         nX1 = cX1 '+ WInProfil1
         nY1 = cY1 '+ -WProfil
         nX2 = cX2 '+ -WInProfil2
         nY2 = cY2 '+ -WProfil
'         'Debug.Print "New Point "
'         'Debug.Print Str(cX1) + "," + Str(cY1) + "," + Str(cX2) + "," + Str(cY2)

         P1.X = cX1
         P1.Y = cY1
         P2.X = cX2
         P2.Y = cY2
'         'Debug.Print "Ðáôçìá Point "
'         'Debug.Print Str(p1.X) + "," + Str(p1.Y) + "," + Str(p2.X) + "," + Str(p2.Y)

         'ÐåñéóôñïöÞ ôïõ ôåìá÷ßïõ óôï 0,0
         'P_1 = Rotate(P1, AngOrig) 'ÏÊ
         P_2 = Rotate(P2.X, P2.Y, AngOrig) 'ÏÊ
'         'Debug.Print "ROTATE Point "
'         'Debug.Print Str(p1.X) + "," + Str(p1.Y) + "," + Str(p2.X) + "," + Str(p2.Y)
         
'         'Ìåôáêßíçóç ôùí óçìåßùí óôá ÔåëéêÜ Óçìåßá
         'P1.X = P_1.X + X1
         'P1.Y = -(P_1.Y + Y1)
         P.X = P_2.X + X1
         P.Y = -(P_2.Y + Y1)
        
         Y1 = -Y1
         Y2 = -Y2
         'ReDim PointDraw(1 To 4) As PointApi
         'PointDraw(1).X = P1.X
         'PointDraw(2).X = p2.X
        
         'PointDraw(1).Y = P1.Y
         'PointDraw(2).Y = p2.Y
         PointInLine = P
         'Debug.Print "MOVE point:" + Str(p.X) + "," + Str(p.Y)
         
End Function


Public Function GetNearestPoint(refx As Long, refy As Long, _
                                endx As Long, endy As Long, _
                                hBoundaryDC As Long, _
                                Radius As Long, DegAngle As Single) As POINTAPI

'finds the first non-white pixel that is closest to (refx, refy) along the line at an angle
'of DegAngle out to a radius of Radius.  traces out a line starting at (refx, refy), checking
'each pixel.

'returns the closest point if there is one, and the endpoint of the radial line (endx, endy)
'if there isn't.

Dim CurrPt As POINTAPI
Dim CurrDistance As Long
Dim StepParam As Long                'used to step along the line

CurrPt.X = refx
CurrPt.Y = refy
CurrDistance = 0
StepParam = 0

While CurrDistance <= Radius

    'check if the current pixel on the boundary map is non-white
    If GetPixel(hBoundaryDC, CurrPt.X, CurrPt.Y) <> RGB(255, 255, 255) Then
        'yes, return current point
        GetNearestPoint = CurrPt
        Exit Function
    Else

        'increment step parameter
        StepParam = StepParam + 1

        'set the next pixel value along the line to check
        CurrPt.X = (refx + StepParam * Cos(PI * DegAngle / 180))
        CurrPt.Y = (refy + StepParam * Sin(PI * DegAngle / 180))

        'calculate the new distance
        CurrDistance = CLng(((CurrPt.X - refx) ^ 2 + (CurrPt.Y - refy) ^ 2) ^ 0.5)
    End If

Wend

'if we get here, then there was no intersection
GetNearestPoint.X = endx
GetNearestPoint.Y = endy

End Function

