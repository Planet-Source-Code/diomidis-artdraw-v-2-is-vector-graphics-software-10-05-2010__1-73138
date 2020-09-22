VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdPolygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(C) 2007-2009 diomidisk
Option Explicit
' VbDraw Polygon/Polyline object.

Implements vbdObject

' Indicates a closed polygon rather than a polyline.
Public IsClosed As Boolean
Public IsBox As Boolean

' Constituent vbdPolygon object.
Private m_Polygon As vbdDraw
Private m_Object As vbdObject

' The surface on which the user is clicking
' to define the object. This is set only during
' creation of this object.
Public WithEvents m_Canvas As PictureBox
Attribute m_Canvas.VB_VarHelpID = -1
Public m_DrawingStarted As Boolean
Public m_DrawingEdit As Boolean
Private m_DrawingMove As Boolean
Private m_SelectPoint As Integer
Private m_SelectPoint2 As Integer 'for nodes

' Drawing properties.
Public m_DrawWidth As Integer
Public m_DrawStyle As Integer
Public m_ForeColor As OLE_COLOR
Public m_FillColor As OLE_COLOR
Public m_FillColor2 As OLE_COLOR
Public m_FillMode As FillMode
Public m_Pattern As String
Public m_Gradient As Integer
Public m_FillStyle As Integer
Public m_TypeDraw As DrawType
Public m_TextDraw As String
Public m_AlingText As Integer
Public m_TypeFill As DrawTypeFill
Public m_Shade As Boolean
Public m_ObjLock As Boolean
Public m_Blend As Integer
'Public m_TextProperty As StdFont
Public m_Bold As Boolean
Public m_Charset As Integer
Public m_Italic As Boolean
Public m_Name As String
Public m_Size As Integer
Public m_Strikethrough As Boolean
Public m_Underline As Boolean
Public m_Weight As Integer
Public m_CurrentX As Single
Public m_CurrentY As Single
Public m_angle As Single
Public m_Selected As Boolean
Public m_hRegion As Long
Public m_Picture As StdPicture

'' Data variables.
Private m_NumPoints As Long
Private m_TypePoint() As Byte
Private m_OriginalPoints() As PointAPI
Private m_TransformedPoints() As PointAPI
Private i As Integer

' Rubberband variables.
Private m_StartX As Single
Private m_StartY As Single
Private m_LastX As Single
Private m_LastY As Single



Private Sub DrawPoint(Optional is_Draw As Boolean = False)
      Dim aa As Integer, mStep As Single
       Dim OldScale As typScaleMode
      
     If m_NumPoints = 0 Then m_Canvas_LostFocus: Exit Sub
      If m_TypePoint(m_NumPoints) = 3 Then
         m_OriginalPoints(m_NumPoints) = m_OriginalPoints(1)
      End If
      
      OldScale = SetWordScale(m_Canvas)
      PolyDraw m_Canvas.hDC, m_OriginalPoints(1), m_TypePoint(1), m_NumPoints
      ResetWordScale m_Canvas, OldScale
            
       mStep = IIf(GAP / gZoomFactor > 0.5, GAP / gZoomFactor, 1)
    
      For i = 1 To m_NumPoints
          If m_SelectPoint <> i Then
              m_Canvas.Line (m_OriginalPoints(i).X - mStep, m_OriginalPoints(i).Y - mStep)-(m_OriginalPoints(i).X + mStep, m_OriginalPoints(i).Y + mStep), vbGreen, BF
          ElseIf m_SelectPoint = i Then
              m_Canvas.Line (m_OriginalPoints(i).X - mStep, m_OriginalPoints(i).Y - mStep)-(m_OriginalPoints(i).X + mStep, m_OriginalPoints(i).Y + mStep), vbRed, BF
          End If
              m_Canvas.Line (m_OriginalPoints(i).X - mStep, m_OriginalPoints(i).Y - mStep)-(m_OriginalPoints(i).X + mStep, m_OriginalPoints(i).Y + mStep), , B
      Next
      If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
         i = m_SelectPoint
         m_Canvas.Line (m_OriginalPoints(i).X - mStep, m_OriginalPoints(i).Y - mStep)-(m_OriginalPoints(i).X + mStep, m_OriginalPoints(i).Y + mStep), vbRed, BF
         m_Canvas.Line (m_OriginalPoints(i).X - mStep, m_OriginalPoints(i).Y - mStep)-(m_OriginalPoints(i).X + mStep, m_OriginalPoints(i).Y + mStep), , B
      End If
      
      aa = 0
      For i = 1 To m_NumPoints
          If aa = 3 Then aa = 0
          If m_TypePoint(i) = 4 Then
             aa = aa + 1
             If aa = 1 Or aa = 3 Then
                 m_Canvas.Line (m_OriginalPoints(i - 1).X, m_OriginalPoints(i - 1).Y)-(m_OriginalPoints(i).X, m_OriginalPoints(i).Y), vbRed
             End If
          End If
       Next
      ' m_Canvas.DrawStyle = oldmode
       
      m_Canvas.Refresh
      If is_Draw = True Then
            Me.m_DrawStyle = Me.vbdObject_DrawStyle
            Me.m_DrawWidth = Me.vbdObject_DrawWidth
            Me.m_FillColor = Me.vbdObject_FillColor
            Me.m_FillColor2 = Me.vbdObject_FillColor2
            Me.m_FillMode = Me.vbdObject_FillMode
            Me.m_FillStyle = Me.vbdObject_FillStyle
            Me.m_ForeColor = Me.vbdObject_ForeColor
            Me.m_TextDraw = Me.vbdObject_TextDraw
            Me.m_TypeDraw = Me.vbdObject_TypeDraw
            Me.m_Blend = Me.vbdObject_TypeFill
            Me.m_Gradient = Me.vbdObject_Gradient
            Me.m_Pattern = Me.vbdObject_Pattern
            Me.m_Shade = Me.vbdObject_Shade
            Me.m_Blend = Me.vbdObject_Blend
            'Me.m_ShadeX=Me.vbdObject_ShadeX
            'Me.m_ShadeY=Me.vbdObject_ShadeY
            'Me.m_ShadeFillStyle=Me.vbdObject_ShadeFillStyle
            'Me.m_ShadeGradient=Me.vbdObject_ShadeGradient
            'Me.m_ShadePattern=Me.vbdObject_ShadePattern
            'Me.m_Gradient  = Me.vbdObject_Gradient
      End If
End Sub

Public Sub PolyPoints(nPoint As Integer, cx As Single, cy As Single)
    Dim isCntl  As Boolean
   ' isCntl = IsControl(m_TypePoint, nPoint)
    If nPoint > 0 Then
        m_OriginalPoints(nPoint).X = cx
        m_OriginalPoints(nPoint).Y = cy
'        If isCntl And nPoint = 1 Then
'           m_OriginalPoints(nPoint).X = cx
'           m_OriginalPoints(nPoint).Y = cy
'        Else
'        End If
    End If
End Sub
'
'Private Sub PolyDrawVB(ByVal hdc As Long, ByRef lpPt() As POINTAPI, _
'    ByRef lpbTypes() As Byte, ByVal cCount As Long)
'    Dim LoopPts As Long
'    Dim BezIdx As Long
'    Command2.Enabled = False
'    For LoopPts = 0 To cCount - 1
'        ' Clear bezier point index for non-bezier point
'        If ((lpbTypes(LoopPts) And PT_BEZIERTO) = 0) Then
'           BezIdx = 0
'        End If
'        Select Case lpbTypes(LoopPts) And Not PT_CLOSEFIGURE
'            Case PT_LINETO    ' Straight line segment
'              '  Call DrawBox(lpPt(LoopPts), 2, vbRed)
'                Me.ForeColor = RGB(200, 50, 70)
'                Call LineTo(hdc, lpPt(LoopPts).X, lpPt(LoopPts).Y)
'                Me.Caption = LoopPts + 1 & "/" & cCount & " LINETO"
'        Case PT_BEZIERTO    ' Curve segment
'                '//
'                Select Case BezIdx
'                    Case 0, 1   ' Bezier control handles
'                        Call DrawBox(lpPt(LoopPts), 2, vbBlue)
'                        Me.Caption = LoopPts + 1 & "/" & cCount & " Bezier control handles :" & BezIdx
'                    Case 2    ' Bezier end point
'                            '//Connecting lines betweenn (start to 1st control pt) and (2nd control pt to end point)
'                            Me.Line (lpPt(LoopPts - 3).X, lpPt(LoopPts - 3).Y)-(lpPt(LoopPts - 2).X, lpPt(LoopPts - 2).Y), vbBlue
'                            Me.Line (lpPt(LoopPts - 1).X, lpPt(LoopPts - 1).Y)-(lpPt(LoopPts).X, lpPt(LoopPts).Y), vbBlue
'                            Me.ForeColor = vbCyan
'                            ' Move to first point where we will start curve
'                            Call MoveToEx(hdc, m_PointCoords(LoopPts - 3).X, m_PointCoords(LoopPts - 3).Y, ByVal 0&)
'                            Call PolyBezierTo(hdc, m_PointCoords(LoopPts - 2), 3)
'                         '   Call DrawDot(lpPt(LoopPts), 2, &HFF00AA)
'                            Me.Caption = LoopPts + 1 & "/" & cCount & " Bezier End Point :" & BezIdx
'                End Select
'                'Debug.Print Me.Caption
'                BezIdx = (BezIdx + 1) Mod 3 '//Reset counter after 3 Bezier points
'            Case PT_MOVETO    ' Move current drawing point
'              '  Call DrawDot(lpPt(LoopPts), 4, RGB(0, 150, 50))
'                Call MoveToEx(hdc, lpPt(LoopPts).X, lpPt(LoopPts).Y, ByVal 0&)
'                Me.Caption = LoopPts + 1 & "/" & cCount & " MOVETO"
'        End Select
'
'        If (lpbTypes(LoopPts) And PT_CLOSEFIGURE) Then
'            Call CloseFigure(hdc)
'          '  Call DrawDot(lpPt(LoopPts), 6, vbYellow)
'        End If
'        Delay (0.01)
'    Next LoopPts
'    Command2.Enabled = True
'End Sub
'
''Check is Control select point
'Private Function IsControl(ByRef lTypes() As Byte, ByVal cCount As Long) As Boolean
'        Dim BezIdx As Long, id As Long
'        Const PT_CLOSEFIGURE As Long = &H1
'        Const PT_LINETO As Long = &H2
'        Const PT_BEZIERTO As Long = &H4
'        Const PT_MOVETO As Long = &H6
'        If cCount = 0 Then Exit Function
'        For id = 1 To cCount
'            If ((lTypes(id) And PT_BEZIERTO) = 0) Then
'               BezIdx = 0
'            End If
'            Select Case lTypes(id) And Not PT_CLOSEFIGURE
'            Case PT_LINETO    ' Straight line segment
'            Case PT_BEZIERTO    ' Curve segment
'                  Select Case BezIdx
'                  Case 0, 1   ' Bezier control handles
'                      IsControl = True
'                  Case 2    ' Bezier end point
'                      IsControl = False
'                  End Select
'                  BezIdx = (BezIdx + 1) Mod 3 '//Reset counter after 3 Bezier points
'            Case PT_MOVETO    ' Move current drawing point
'            End Select
'        Next
'End Function
'
'' Check for opening
'Private Function IsOpening(ByRef lTypes() As Byte) As Boolean
'        Dim cCount As Long, id As Long, aa As Long
'
'        cCount = UBound(lTypes)
'        If cCount = 0 Then Exit Function
'        For id = 1 To cCount
'            If lTypes(id) = 3 Then
'               aa = aa + 1
'            End If
'        Next
'        'aa = aa - 1
'        If aa > 0 Then IsOpening = True Else IsOpening = False
'End Function
'
'' Calculate new Point in the line.
'Private Function AddNode(ByVal X As Single, ByVal Y As Single) As POINTAPI
'
'        Dim MinX As Single, MinY As Single, I As Long, e As Long, nD As Integer
'        Dim NewDist As Single, mDist As Single, aa As Long, t As Long
'        Dim X1 As Single, Y1 As Single, X2 As Single, Y2 As Single
'        Dim Points() As POINTAPI, mTypePoint() As Byte
'
'        'add sto telos
'        NewDist = 0
'        e = 0
'        For I = 1 To m_NumPoints - 1
'            mDist = DistToSegment(X, Y, m_OriginalPoints(I).X, m_OriginalPoints(I).Y, _
'                                        m_OriginalPoints(I + 1).X, m_OriginalPoints(I + 1).Y, MinX, MinY)
'            If NewDist >= mDist Or NewDist = 0 Then
'               NewDist = mDist
'               e = I + 1
'               'if On the node then find midpoint from next node
'               If MinX + mStep  >= m_OriginalPoints(I).X And MinX - mStep  <= m_OriginalPoints(I).X Then
'               If MinY + mStep  >= m_OriginalPoints(I).Y And MinY - mStep  <= m_OriginalPoints(I).Y Then
'                  MidPoint m_OriginalPoints(I).X, m_OriginalPoints(I).Y, _
'                           m_OriginalPoints(I + 1).X, m_OriginalPoints(I + 1).Y, _
'                           MinX, MinY
'               End If
'               End If
'               If m_TypePoint(e) = 4 Then
'                  MidPoint m_OriginalPoints(I).X, m_OriginalPoints(I).Y, _
'                           MinX, MinY, _
'                           X1, Y1
'                  MidPoint MinX, MinY, _
'                           m_OriginalPoints(I + 1).X, m_OriginalPoints(I + 1).Y, _
'                           X2, Y2
'               Else
'                  X1 = MinX
'                  Y1 = MinY
'                  X2 = MinX
'                  Y2 = MinY
'               End If
'               AddNode.X = MinX
'               AddNode.Y = MinY
'            End If
'        Next
'
'         If e >= 0 And e <= m_NumPoints Then
'             'm_Canvas.Circle (AddNode.X, AddNode.Y), 5
'             'Check Curver
'             If m_TypePoint(e) = 4 Then nD = 3 Else nD = 1
'
'             ReDim Points(1 To m_NumPoints + nD)
'             ReDim mTypePoint(1 To m_NumPoints + nD)
'             aa = 0
'             For I = 1 To m_NumPoints + nD
'                If e = I And nD = 1 Then
'                  Points(I).X = AddNode.X
'                  Points(I).Y = AddNode.Y
'                  mTypePoint(I) = 2
'
'                ElseIf e = I And nD = 3 Then
'                  Points(I).X = X1 'AddNode.X + 10
'                  Points(I).Y = Y1 'AddNode.Y
'                  mTypePoint(I) = 4
'
'                  Points(I + 1).X = AddNode.X
'                  Points(I + 1).Y = AddNode.Y
'                  mTypePoint(I + 1) = 4
'
'                  Points(I + 2).X = X2 'AddNode.X - 10
'                  Points(I + 2).Y = Y2 'AddNode.Y
'                  mTypePoint(I + 2) = 4
'                  I = I + 2
'                Else
'                    aa = aa + 1
'                    Points(I).X = m_OriginalPoints(aa).X
'                    Points(I).Y = m_OriginalPoints(aa).Y
'                    mTypePoint(I) = m_TypePoint(aa)
'                End If
'             Next
'             m_NumPoints = m_NumPoints + nD
'             m_OriginalPoints = Points
'             m_TypePoint = mTypePoint
'        End If
'
'End Function
'
''Delete select node
'Private Sub DeleteNode()
'        Dim Points() As POINTAPI, aa As Long, I As Long, t As Integer
'        Dim mTypePoint() As Byte, sPoint As Integer, ePoint As Integer
'        Dim Arr()
'        ReDim Arr(0)
'        If m_NumPoints > 2 Then
'           If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
'                If m_SelectPoint + 1 = m_NumPoints Then
'                   If m_TypePoint(m_SelectPoint + 1) = 2 Then
'                      m_NumPoints = m_NumPoints - 1
'                      ReDim Points(1 To m_NumPoints)
'                      ReDim mTypePoint(1 To m_NumPoints)
'                      For I = 1 To m_NumPoints
'                         Points(I) = m_OriginalPoints(I)
'                          mTypePoint(I) = m_TypePoint(I)
'                      Next
'                      m_OriginalPoints = Points
'                      m_TypePoint = mTypePoint
'                   End If
'                ElseIf m_TypePoint(m_SelectPoint) = 4 Then
'                    If IsControl(m_TypePoint, m_SelectPoint) = False Then
'                        For I = 1 To m_NumPoints '- 1
'                            For t = I To I + 3
'                                If m_TypePoint(I) = 4 Then
'                                    If Arr(UBound(Arr)) <> I + 2 Then
'                                        ReDim Preserve Arr(UBound(Arr) + 1)
'                                        Arr(UBound(Arr)) = I + 2 ' m_TypePoint(I)
'                                        If Arr(UBound(Arr)) > m_NumPoints Then
'                                            Arr(UBound(Arr)) = m_NumPoints
'                                        End If
'                                        t = t + 3
'                                        I = I + 2
'                                        Exit For
'                                    End If
'                                Else
'                                    If Arr(UBound(Arr)) <> I Then
'                                        ReDim Preserve Arr(UBound(Arr) + 1)
'                                        Arr(UBound(Arr)) = I 'm_TypePoint(I)
'                                        End If
'                                    End If
'                            Next
'                        Next
'                        For I = 1 To UBound(Arr) - 1
'                            If m_SelectPoint >= Arr(I) And m_SelectPoint < Arr(I + 1) Then
'                                sPoint = I
'                            End If
'                        Next
'                        If sPoint > 0 Then ePoint = sPoint + 1 Else Exit Sub
'
'                        m_NumPoints = m_NumPoints - (Arr(ePoint) - Arr(sPoint))
'
'                        ReDim Points(1 To m_NumPoints)
'                        ReDim mTypePoint(1 To m_NumPoints)
'                        aa = 0
'                        For I = 1 To UBound(m_OriginalPoints) '+ 1
'                            If I >= Arr(sPoint) And I < Arr(ePoint) Then
'                            'Stop
'                            Else
'                                aa = aa + 1
'                                Points(aa).X = m_OriginalPoints(I).X
'                                Points(aa).Y = m_OriginalPoints(I).Y
'                                mTypePoint(aa) = m_TypePoint(I)
'                            End If
'                        Next
'                        If m_SelectPoint = UBound(m_TypePoint) Then mTypePoint(1) = m_TypePoint(UBound(m_TypePoint))
'                        If m_SelectPoint = 1 Or mTypePoint(1) <> 6 Then mTypePoint(1) = 6
'                        m_OriginalPoints = Points
'                        m_TypePoint = mTypePoint
'                        ''Debug.Print sPoint, sPoint + 2, ePoint
'                        ' Stop
'                    End If
'                Else
'                    m_NumPoints = m_NumPoints - 1
'                    ReDim Points(1 To m_NumPoints)
'                    ReDim mTypePoint(1 To m_NumPoints)
'                    aa = 0
'                    For I = 1 To m_NumPoints + 1
'                        If m_SelectPoint <> I Then
'                            aa = aa + 1
'                            Points(aa).X = m_OriginalPoints(I).X
'                            Points(aa).Y = m_OriginalPoints(I).Y
'                            mTypePoint(aa) = m_TypePoint(I)
'                        End If
'                    Next
'                    If m_SelectPoint = UBound(m_TypePoint) Then mTypePoint(1) = m_TypePoint(UBound(m_TypePoint))
'                    If m_SelectPoint = 1 Or mTypePoint(1) <> 6 Then mTypePoint(1) = 6
'                    m_OriginalPoints = Points
'                    m_TypePoint = mTypePoint
'                End If
'           End If
'           DrawPoint
'        Else
'           m_NumPoints = 0
'        End If
'End Sub
'
''Make Curve to Line
'Private Sub toLine()
'        Dim Points() As POINTAPI, p() As POINTAPI, aa As Long, I As Long, t As Integer
'        Dim mTypePoint() As Byte, sPoint As Integer, ePoint As Integer
'        Dim Arr()
'        ReDim Arr(0)
'        If m_NumPoints > 2 Then
'          If IsControl(m_TypePoint, m_SelectPoint) Then Exit Sub
'          m_SelectPoint = m_SelectPoint + 1
'
'          If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
'             If m_SelectPoint > 1 Then m_SelectPoint = m_SelectPoint + 1
'             If m_SelectPoint >= m_NumPoints Then m_SelectPoint = m_NumPoints - 1
'               For I = 1 To m_NumPoints - 1
'                     For t = I To I + 3
'                      If m_TypePoint(I) = 4 Then
'                         If Arr(UBound(Arr)) <> I + 2 Then
'                         ReDim Preserve Arr(UBound(Arr) + 1)
'                         Arr(UBound(Arr)) = I + 2
'                         If Arr(UBound(Arr)) > m_NumPoints Then
'                            Arr(UBound(Arr)) = m_NumPoints
'                         End If
'                         I = I + 2
'                         Exit For
'                         End If
'                      Else
'                        If Arr(UBound(Arr)) <> I Then
'                        ReDim Preserve Arr(UBound(Arr) + 1)
'                        Arr(UBound(Arr)) = I
'                        End If
'                     End If
'                    Next
'               Next
'               For I = 1 To UBound(Arr) - 1
'                   If m_SelectPoint >= Arr(I) And m_SelectPoint < Arr(I + 1) Then
'                      sPoint = I
'                   End If
'                Next
'                'If sPoint > 0 Then ePoint = sPoint + 1
'                If sPoint > 0 Then ePoint = sPoint + 1 Else Exit Sub
'                m_NumPoints = m_NumPoints - (((Arr(ePoint)) - (Arr(sPoint))) - 1)
'
'                p = m_OriginalPoints
'                For I = Arr(sPoint) + 1 To Arr(ePoint) - 1
'                    p(I).X = 0
'                    p(I).Y = 0
'                Next
'
'                ReDim Points(1 To m_NumPoints)
'                ReDim mTypePoint(1 To m_NumPoints)
'                aa = 0
'                For I = 1 To UBound(p)
'                  If p(I).X <> 0 Then
'                    aa = aa + 1
'                    Points(aa).X = m_OriginalPoints(I).X
'                    Points(aa).Y = m_OriginalPoints(I).Y
'                    mTypePoint(aa) = m_TypePoint(I)
'                 End If
'                Next
'                mTypePoint(Arr(sPoint) + 1) = 2
'                m_OriginalPoints = Points
'                m_TypePoint = mTypePoint
'          End If
'          DrawPoint
'        End If
'End Sub
'
''Make Line to Curve
'Private Sub toCurve()
'     Dim Points() As POINTAPI, aa As Long, mType As Byte, I As Long
'     Dim mTypePoint() As Byte, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single
'     Dim m_EndPoint As Boolean
'     If m_NumPoints > 2 Then
'         If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
'             If m_NumPoints = m_SelectPoint + 1 Then m_EndPoint = True
'             ReDim Points(1 To m_NumPoints + 2)
'             ReDim mTypePoint(1 To m_NumPoints + 2)
'             m_SelectPoint = m_SelectPoint + 1
'             If m_SelectPoint = 1 Then m_SelectPoint = 2
'             MidPoint m_OriginalPoints(m_SelectPoint).X, m_OriginalPoints(m_SelectPoint).Y, _
'                      m_OriginalPoints(m_SelectPoint - 1).X, m_OriginalPoints(m_SelectPoint - 1).Y, _
'                      X1, Y1
'             MidPoint m_OriginalPoints(m_SelectPoint).X, m_OriginalPoints(m_SelectPoint).Y, X1, Y1, X2, Y2
'             MidPoint X1, Y1, m_OriginalPoints(m_SelectPoint - 1).X, m_OriginalPoints(m_SelectPoint - 1).Y, X3, Y3
'             aa = 0
'             For I = 1 To m_SelectPoint + 2
'                If m_SelectPoint >= I - 2 And m_SelectPoint <= I Then
'                    aa = aa + 1
'                    Points(aa).X = m_OriginalPoints(m_SelectPoint).X
'                    Points(aa).Y = m_OriginalPoints(m_SelectPoint).Y
'                    mTypePoint(aa) = 4
'                Else
'                    aa = aa + 1
'                    Points(aa).X = m_OriginalPoints(I).X
'                    Points(aa).Y = m_OriginalPoints(I).Y
'                    mTypePoint(aa) = m_TypePoint(I)
'                End If
'             Next
'             For I = m_SelectPoint + 1 To m_NumPoints
'                  aa = aa + 1
'                  Points(aa).X = m_OriginalPoints(I).X
'                  Points(aa).Y = m_OriginalPoints(I).Y
'                  mTypePoint(aa) = m_TypePoint(I)
'             Next
'             Points(m_SelectPoint).X = X3
'             Points(m_SelectPoint).Y = Y3
'             Points(m_SelectPoint + 1).X = X2
'             Points(m_SelectPoint + 1).Y = Y2
'            ' Stop
'
'             If m_TypePoint(m_NumPoints) = 3 Then
'                ReDim Preserve Points(1 To UBound(Points))
'                ReDim Preserve mTypePoint(1 To UBound(mTypePoint))
'                Points(UBound(Points)) = m_OriginalPoints(1)
'                mTypePoint(UBound(mTypePoint)) = 3
'                'm_NumPoints = m_NumPoints + 1
'             End If
'             m_NumPoints = m_NumPoints + 2
'             If m_EndPoint Then
'                m_NumPoints = m_NumPoints + 1
'                ReDim Preserve Points(1 To m_NumPoints)
'                ReDim Preserve mTypePoint(1 To m_NumPoints)
'                Points(m_SelectPoint + 2).X = m_OriginalPoints(UBound(m_OriginalPoints)).X
'                Points(m_SelectPoint + 2).Y = m_OriginalPoints(UBound(m_OriginalPoints)).Y
'                Points(m_NumPoints).X = m_OriginalPoints(UBound(m_OriginalPoints)).X
'                Points(m_NumPoints).Y = m_OriginalPoints(UBound(m_OriginalPoints)).Y
'                mTypePoint(m_NumPoints - 1) = 4
'                mTypePoint(m_NumPoints) = 3
'             End If
'             m_OriginalPoints = Points
'             m_TypePoint = mTypePoint
'         End If
'     End If
'
'End Sub
'
''Close Node line
'Private Sub CloseNode()
'     Dim Points() As POINTAPI, aa As Long, mType As Byte, I As Long
'     Dim mTypePoint() As Byte
'
'     If m_NumPoints >= 2 Then
'         If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
'           ReDim Points(1 To m_NumPoints)
'           ReDim mTypePoint(1 To m_NumPoints)
'           Points = m_OriginalPoints
'           mTypePoint = m_TypePoint
'           For I = 2 To m_NumPoints - 1
'              If mTypePoint(I) = 6 Or mTypePoint(I) = 3 Then
'                 mTypePoint(I) = 2
'              ElseIf mTypePoint(I) = 5 Then
'                 mTypePoint(I) = 4
'              End If
'           Next
''            If mTypePoint(m_NumPoints) <> 4 Then
''               mTypePoint(m_NumPoints) = 3
''            Else
'               If mTypePoint(m_NumPoints) <> 3 Then
'                 m_NumPoints = m_NumPoints + 1
'                 ReDim Preserve Points(1 To m_NumPoints)
'                 ReDim Preserve mTypePoint(1 To m_NumPoints)
'                 mTypePoint(m_NumPoints) = 3
'                 Points(m_NumPoints) = m_OriginalPoints(1)
'               End If
''            End If
'          m_OriginalPoints = Points
'          m_TypePoint = mTypePoint
'          DrawPoint
'         End If
'     End If
'End Sub
'
''Open Node line
'Private Sub OpenNode()
'     Dim Points() As POINTAPI, aa As Long, mType As Byte, I As Long
'     Dim mTypePoint() As Byte
'
'     If m_NumPoints >= 2 Then
'         If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
'           ReDim Points(1 To m_NumPoints)
'           ReDim mTypePoint(1 To m_NumPoints)
'           Points = m_OriginalPoints
'           mTypePoint = m_TypePoint
'           For I = 2 To m_NumPoints - 1
'              If mTypePoint(I) = 3 Then
'                 mTypePoint(I) = 2
'              End If
'           Next
'          If mTypePoint(I) = 3 Then
'             m_NumPoints = m_NumPoints - 1
'             ReDim Preserve Points(1 To m_NumPoints)
'             ReDim Preserve mTypePoint(1 To m_NumPoints)
'          End If
'          m_OriginalPoints = Points
'          m_TypePoint = mTypePoint
'          DrawPoint
'         End If
'     End If
'End Sub
'
''Break line in select node
'Private Sub BreakNode()
'    Dim Points() As POINTAPI, aa As Long, mType As Byte, I As Long
'    Dim mTypePoint() As Byte, np As Integer, t As Long
'    Dim Arr()
'        ReDim Arr(0)
'
'       If m_NumPoints >= 2 Then
'         If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
'           If m_SelectPoint = 1 Or m_SelectPoint = m_NumPoints Then
'
'              ReDim Points(1 To m_NumPoints)
'              ReDim mTypePoint(1 To m_NumPoints)
'              Points = m_OriginalPoints
'              mTypePoint = m_TypePoint
'              If m_TypePoint(m_NumPoints) = 3 Then
'                 mType = 2
'                 mTypePoint(m_NumPoints) = 2
'              ElseIf m_TypePoint(m_NumPoints) = 4 Then
'                 mType = 2
'                 mTypePoint(m_NumPoints) = 4
'              Else
'                 mType = 2
'              End If
'              m_NumPoints = m_NumPoints + 1
'              ReDim Preserve Points(1 To m_NumPoints)
'              ReDim Preserve mTypePoint(1 To m_NumPoints)
'              Points(m_NumPoints) = m_OriginalPoints(m_SelectPoint)
'              mTypePoint(m_NumPoints) = mType
'           Else
'              m_NumPoints = m_NumPoints + 1
'              ReDim Points(1 To m_NumPoints)
'              ReDim mTypePoint(1 To m_NumPoints)
'               aa = 0
'               aa = aa + 1
'               Points(aa).X = m_OriginalPoints(1).X
'               Points(aa).Y = m_OriginalPoints(1).Y
'               mTypePoint(aa) = m_TypePoint(1)
''               If m_TypePoint(m_SelectPoint) = 2 Then
''                  np = 1
''               End If
'                 For I = 2 To m_SelectPoint - np
'
'                  aa = aa + 1
'                  Points(aa).X = m_OriginalPoints(I).X
'                  Points(aa).Y = m_OriginalPoints(I).Y
'                  mTypePoint(aa) = m_TypePoint(I)
'               Next
'               aa = aa + 1
'               Points(aa).X = m_OriginalPoints(m_SelectPoint).X
'               Points(aa).Y = m_OriginalPoints(m_SelectPoint).Y
'               mTypePoint(aa) = 6
''               aa = aa + 1
''               Points(aa).X = m_OriginalPoints(m_SelectPoint).X
''               Points(aa).Y = m_OriginalPoints(m_SelectPoint).Y
''              ' If m_TypePoint(m_SelectPoint) <> 4 Then
''                  mTypePoint(aa) = m_TypePoint(m_SelectPoint)
''              ' Else
''              '   mTypePoint(aa) = 6
''              ' End If
'               For I = m_SelectPoint + 1 To m_NumPoints - 1 '- 3
'                  aa = aa + 1
'                  Points(aa).X = m_OriginalPoints(I).X
'                  Points(aa).Y = m_OriginalPoints(I).Y
'                  mTypePoint(aa) = m_TypePoint(I)
'               Next
'               aa = 0
'               For I = 1 To m_NumPoints
'                  If mTypePoint(I) = 6 Then aa = aa + 1
'               Next
'               If aa > 1 Then
'                  If mTypePoint(m_NumPoints) = 3 Then
'                     mTypePoint(m_NumPoints) = 2
'                     m_NumPoints = m_NumPoints + 1
'                     ReDim Preserve Points(1 To m_NumPoints)
'                     ReDim Preserve mTypePoint(1 To m_NumPoints)
'                     Points(m_NumPoints) = m_OriginalPoints(1)
'                     mTypePoint(m_NumPoints) = 2
'                   ElseIf mTypePoint(m_NumPoints) = 0 Then
'                      mTypePoint(m_NumPoints) = m_TypePoint(UBound(m_TypePoint))
'                  End If
'               End If
'
'            End If
'             m_OriginalPoints = Points
'             m_TypePoint = mTypePoint
'          End If
'          DrawPoint
'        End If
'End Sub
''
'Private Sub FindNodeCurve(ByVal px1 As Single, ByVal py1 As Single, ByVal px2 As Single, ByVal py2 As Single, _
'                          ByRef cX1 As Single, ByRef cY1 As Single, ByRef cX2 As Single, ByRef cY2 As Single)
'            Dim tX1 As Single, tY1 As Single
'
'            MidPoint px1, py1, px2, py2, tX1, tY1
'            MidPoint px1, py1, tX1, tY1, cX1, cY1
'            MidPoint tX1, tY1, px2, py2, cX2, cY2
'End Sub
'
''Popup Menu
'Private Function MenuNode() As Long
'    Dim Pt As POINTAPI
'    Dim ret As Long
'    Dim wFlag0 As Long, wFlag1 As Long, wFlag2 As Long, wFlag3 As Long, wFlag4 As Long, wFlag5 As Long, wFlag6 As Long
'
'    If IsControl(m_TypePoint, m_SelectPoint) = True Then
'       wFlag0 = MF_GRAYED Or MF_DISABLED
'    Else
'        wFlag0 = MF_STRING
'    End If
'    If m_SelectPoint = 0 Or m_SelectPoint = 1 Or IsControl(m_TypePoint, m_SelectPoint) = True Then
'        wFlag1 = MF_GRAYED Or MF_DISABLED
'    Else
'        wFlag1 = MF_STRING
'    End If
'
'    If m_SelectPoint > 0 Then
'       If m_TypePoint(m_SelectPoint) = 6 Then
'          If m_SelectPoint + 1 > m_NumPoints Then
'               wFlag2 = MF_GRAYED Or MF_DISABLED
'              wFlag3 = MF_GRAYED Or MF_DISABLED
'          Else
'          If m_TypePoint(m_SelectPoint + 1) = 2 Then wFlag2 = MF_GRAYED Or MF_DISABLED Else wFlag2 = MF_STRING
'          If m_TypePoint(m_SelectPoint + 1) = 4 Then wFlag3 = MF_GRAYED Or MF_DISABLED Else wFlag3 = MF_STRING
'          End If
'       ElseIf m_TypePoint(m_SelectPoint) = 3 Then
'           wFlag2 = MF_GRAYED Or MF_DISABLED
'           wFlag3 = MF_GRAYED Or MF_DISABLED
'       Else
'          If m_SelectPoint >= m_NumPoints Then m_SelectPoint = m_SelectPoint - 1
'          If m_TypePoint(m_SelectPoint + 1) = 2 Or _
'             m_TypePoint(m_SelectPoint + 1) = 6 Or _
'             m_TypePoint(m_SelectPoint + 1) = 3 Or _
'             IsControl(m_TypePoint, m_SelectPoint) = True Then
'              wFlag2 = MF_GRAYED Or MF_DISABLED
'          Else
'              wFlag2 = MF_STRING
'          End If
'
'          If m_TypePoint(m_SelectPoint + 1) = 4 Or _
'             m_TypePoint(m_SelectPoint + 1) = 6 Or _
'             m_TypePoint(m_SelectPoint) = 3 Or _
'             IsControl(m_TypePoint, m_SelectPoint) = True Then
'             wFlag3 = MF_GRAYED Or MF_DISABLED
'          Else
'             wFlag3 = MF_STRING
'          End If
'
'       End If
'    Else
'       wFlag2 = MF_GRAYED Or MF_DISABLED
'       wFlag3 = MF_GRAYED Or MF_DISABLED
'    End If
'
'    If m_TypePoint(m_NumPoints) = 3 Or m_TypePoint(m_NumPoints) = 5 Then wFlag4 = MF_GRAYED Or MF_DISABLED Else wFlag4 = MF_STRING
'
'    If IsOpening(m_TypePoint) Then wFlag5 = MF_STRING Else wFlag5 = MF_GRAYED Or MF_DISABLED
'
'    If m_SelectPoint > 0 And m_SelectPoint <> m_NumPoints And m_TypePoint(m_NumPoints) <> 6 And _
'       IsControl(m_TypePoint, m_SelectPoint) = False Then
'       wFlag6 = MF_STRING
'    Else
'       wFlag6 = MF_GRAYED Or MF_DISABLED
'    End If
'
'    hMenu = CreatePopupMenu()
'    AppendMenu hMenu, wFlag0, 1, "Add node(s)"
'    AppendMenu hMenu, wFlag1, 2, "Delete node(s)" + vbTab + "(Del)"
'    AppendMenu hMenu, MF_SEPARATOR, 3, ByVal 0&
'    AppendMenu hMenu, wFlag2, 4, "To Line"
'    AppendMenu hMenu, wFlag3, 5, "To Curve"
'    AppendMenu hMenu, MF_SEPARATOR, 6, ByVal 0&
'    AppendMenu hMenu, wFlag4, 7, "Auto Close"
'    AppendMenu hMenu, wFlag5, 8, "Auto Open"
'    AppendMenu hMenu, MF_SEPARATOR, 9, ByVal 0&
'    AppendMenu hMenu, wFlag6, 10, "Break node"
'    AppendMenu hMenu, MF_GRAYED Or MF_DISABLED, 11, "Break Apart"
'
'    ''Debug.Print m_TypePoint(m_NumPoints)
'    GetCursorPos Pt
'    ret = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, Pt.X, Pt.Y, m_Canvas.hWnd, ByVal 0&)
'    DestroyMenu hMenu
'   ''Debug.Print ret
'    MenuNode = ret
'End Function
'

Private Sub m_Canvas_KeyUp(KeyCode As Integer, Shift As Integer)
       If KeyCode = vbKeyDelete Then
'          DeleteNode
'          Exit Sub
       ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then

'          m_SelectPoint = m_SelectPoint - 1
'          If m_SelectPoint >= m_NumPoints Then m_SelectPoint = m_NumPoints
'          DrawPoint
       ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
          'If m_SelectPoint = m_NumPoints Then
          '   m_SelectPoint = 1
          'End If
          m_SelectPoint = m_SelectPoint + 1
          If m_SelectPoint > m_NumPoints Then m_SelectPoint = 1
          DrawPoint
       End If
       
End Sub

Private Sub m_Canvas_LostFocus()
  Dim MX As Single, mY As Single
      
       Set m_Polygon = New vbdDraw
       Set m_Object = m_Polygon
       m_DrawingEdit = False
        m_Canvas.DrawMode = vbCopyPen

        ' Stop receiving events from the canvas.
        Set m_Canvas = Nothing

        ' If we have at least 2 points, tell the form to save us.
        If m_NumPoints >= 2 Then
            m_TypePoint(1) = 6
            If IsClosed = True Then
              If m_TypePoint(m_NumPoints) <> 3 Then
               m_NumPoints = m_NumPoints + 1
               ReDim Preserve m_TypePoint(1 To m_NumPoints)
               ReDim Preserve m_OriginalPoints(1 To m_NumPoints)
               m_TypePoint(m_NumPoints) = 3
               m_OriginalPoints(m_NumPoints) = m_OriginalPoints(1)
               End If
            End If
            
            With m_Polygon
                .NumPoints = m_NumPoints
                For i = 1 To m_NumPoints
                     MX = m_OriginalPoints(i).X
                     mY = m_OriginalPoints(i).Y
                  ' PolyPoints I, mX, mY
                   .X(i) = m_OriginalPoints(i).X '/ gZoomFactor
                   .Y(i) = m_OriginalPoints(i).Y '/ gZoomFactor
                   .TypePoint(i) = m_TypePoint(i)
                Next
                .IsClosed = IsClosed
            End With

            Me.vbdObject_DrawStyle = m_DrawStyle
            Me.vbdObject_DrawWidth = m_DrawWidth
            Me.vbdObject_FillColor = m_FillColor
            Me.vbdObject_FillColor2 = m_FillColor2
            Me.vbdObject_FillMode = m_FillMode
            Me.vbdObject_FillStyle = m_FillStyle
            Me.vbdObject_ForeColor = m_ForeColor
            Me.vbdObject_TextDraw = m_TextDraw
            Me.vbdObject_TypeDraw = dPolydraw 'm_TypeDraw
            m_TypeDraw = dPolydraw
            Me.vbdObject_TypeFill = m_TypeFill
            Me.vbdObject_Gradient = m_Gradient
            Me.vbdObject_Pattern = m_Pattern
            Me.vbdObject_Blend = m_Blend
            Me.vbdObject_Shade = m_Shade
            
            Me.vbdObject_Bold = m_Bold
            Me.vbdObject_Charset = m_Charset
            Me.vbdObject_Italic = m_Italic
            Me.vbdObject_Name = m_Name
            Me.vbdObject_Size = m_Size
            Me.vbdObject_Strikethrough = m_Strikethrough
            Me.vbdObject_Underline = m_Underline
            Me.vbdObject_Weight = m_Weight
            
           ' Me.vbdObject_ShadeX = m_ShadeX
           ' Me.vbdObject_ShadeY = m_ShadeY
           ' Me.vbdObject_ShadeFillStyle = m_ShadeFillStyle
           ' Me.vbdObject_ShadeGradient = m_ShadeGradient
           ' Me.vbdObject_ShadePattern = m_ShadePattern
          
            Me.vbdObject_MakeTransformation
            'We have at least 2 points. Tell the form to save us.
            
            AddObject Me

     End If
End Sub

' The user has selected a point.
Private Sub m_Canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Dim SelectMenu As Long, cX1 As Single, cY1 As Single, cX2 As Single, cY2 As Single

    ' If this is the first point, start using dotted vbInvert mode.
    If m_NumPoints = 0 Then
        m_Canvas.DrawMode = vbInvert
        m_Canvas.DrawStyle = vbDot
        m_DrawingStarted = True
    End If
       
    If m_DrawingEdit = False Then
        If m_NumPoints > 0 Then m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
       ' See if this is the left or right button.
       If Button = vbLeftButton Then
            ' It's the left button.
            ' Add this point to the polygon.

            ' Add the new point.
            If m_TypeDraw = dFreePolygon Or m_TypeDraw = dPolyline Or m_TypeDraw = dPolydraw Then
               
               m_NumPoints = m_NumPoints + 1
               ReDim Preserve m_OriginalPoints(1 To m_NumPoints)
               ReDim Preserve m_TypePoint(1 To m_NumPoints)
               
               m_OriginalPoints(m_NumPoints).X = X
               m_OriginalPoints(m_NumPoints).Y = Y
               If m_NumPoints = 1 Then
               m_TypePoint(m_NumPoints) = 6
               Else
               m_TypePoint(m_NumPoints) = 2
               End If
            ElseIf m_TypeDraw = dCurve Then
               
               If m_NumPoints = 0 Then
                  m_NumPoints = 1
                  ReDim Preserve m_OriginalPoints(1 To m_NumPoints)
                  ReDim Preserve m_TypePoint(1 To m_NumPoints)
                  m_OriginalPoints(1).X = X
                  m_OriginalPoints(1).Y = Y
               Else
                  m_NumPoints = m_NumPoints + 3
                  ReDim Preserve m_OriginalPoints(1 To m_NumPoints)
                  ReDim Preserve m_TypePoint(1 To m_NumPoints)
                  FindNodeCurve m_OriginalPoints(m_NumPoints - 3).X, m_OriginalPoints(m_NumPoints - 3).Y, _
                                X, Y, cX1, cY1, cX2, cY2
                  m_OriginalPoints(m_NumPoints - 2).X = cX1
                  m_OriginalPoints(m_NumPoints - 2).Y = cY1
                  m_OriginalPoints(m_NumPoints - 1).X = cX2
                  m_OriginalPoints(m_NumPoints - 1).Y = cY2
                  m_OriginalPoints(m_NumPoints).X = X
                  m_OriginalPoints(m_NumPoints).Y = Y
                  m_TypePoint(m_NumPoints - 2) = 4
                  m_TypePoint(m_NumPoints - 1) = 4
                  m_TypePoint(m_NumPoints) = 4
               
               End If
               DrawPoint
            End If
            ' Draw the line permanently.
            m_Canvas.DrawMode = vbCopyPen
            m_Canvas.DrawStyle = vbSolid
            m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
            m_Canvas.DrawMode = vbInvert
            m_Canvas.DrawStyle = vbDot

            ' Start the next rubberband line.
            m_StartX = X
            m_StartY = Y
            m_LastX = X
            m_LastY = Y
            m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
            
       Else 'end draw
           m_Canvas.DrawMode = vbCopyPen
           m_Canvas.DrawStyle = vbSolid
           DrawPoint
           m_DrawingStarted = False
           m_Canvas_LostFocus
           m_DrawingEdit = False 'True
       End If
       
    Else 'move Node(point)

'        m_SelectPoint = 0
'        m_Canvas.ToolTipText = ""
'        For I = 1 To m_NumPoints
'            If X >= m_OriginalPoints(I).X - mStep  And X <= m_OriginalPoints(I).X + mStep  And _
'               Y >= m_OriginalPoints(I).Y - mStep  And Y <= m_OriginalPoints(I).Y + mStep  Then
'               m_SelectPoint = I
'               'If m_SelectPoint = m_NumPoints Then m_SelectPoint = m_SelectPoint - 1
'               m_Canvas.ToolTipText = "Select Point:" + Str(m_SelectPoint)
'               Exit For
'            End If
'        Next
'
'        If Button = 2 And m_TypeDraw = dPolydraw Then
'            SelectMenu = MenuNode()
'            If SelectMenu = 1 Then  'Add Node
'                AddNode X, Y
'            ElseIf SelectMenu = 2 Then  'Delete Node
'                DeleteNode
'            ElseIf SelectMenu = 4 Then 'Line
'                toLine
'            ElseIf SelectMenu = 5 Then 'Curve
'                toCurve
'            ElseIf SelectMenu = 7 Then 'Auto Close
'                CloseNode
'                IsClosed = True
'            ElseIf SelectMenu = 8 Then 'Auto open
'                OpenNode
'                IsClosed = False
'            ElseIf SelectMenu = 10 Then 'Break
'                BreakNode
'                IsClosed = False
'            End If
'            ' Go back to vbCopyPen drawing mode.
'            m_Canvas.DrawMode = vbCopyPen
'            m_Canvas.DrawStyle = vbSolid
'            m_Canvas_Paint
'            DrawPoint
'            m_DrawingMove = False
'            Exit Sub
'         End If
'
'        If m_SelectPoint > 0 Then
'           m_Canvas.DrawMode = vbInvert
'           m_Canvas.DrawStyle = vbDot
'
'           m_LastX = X
'           m_LastY = Y
'
'           DrawPoint
'           m_DrawingMove = True
'        Else
'           'End Draw
'           m_Canvas_LostFocus
'        End If
    End If
    Exit Sub
End Sub

' Continue drawing the rubberband line.
Private Sub m_Canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'If Not m_DrawingStarted Then Exit Sub
    If m_DrawingStarted Then

        ' Start using dotted vbInvert mode.
        m_Canvas.DrawMode = vbInvert
        m_Canvas.DrawStyle = vbDot

        ' Erase the old line.
        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)

        ' Update the point.
        m_LastX = X
        m_LastY = Y

        ' Draw the new line.
        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
    
'    ElseIf m_DrawingEdit Then
'        If m_DrawingMove = False Then Exit Sub
'
'        DrawPoint
'        ' Update the point.
'        m_LastX = X
'        m_LastY = Y
'
'        Call PolyPoints(m_SelectPoint, m_LastX, m_LastY)
'        DrawPoint
              
   End If
   
End Sub


Private Sub m_Canvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If m_DrawingEdit Then
     
        If m_DrawingMove = False Then Exit Sub
         
         DrawPoint
        
        ' Go back to vbCopyPen drawing mode.
         m_Canvas.DrawMode = vbCopyPen
         m_Canvas.DrawStyle = vbSolid
         m_Canvas_Paint
         
           Call PolyPoints(m_SelectPoint, X, Y)
        
         DrawPoint
        
         m_DrawingMove = False
     End If
End Sub

Private Sub m_Canvas_Paint()
    Dim olddrawmode As Integer
     m_Canvas.Cls
     olddrawmode = m_Canvas.DrawMode
     m_Canvas.DrawMode = 13
     m_TheScene.Draw m_Canvas
     m_Canvas.DrawMode = olddrawmode
End Sub

Private Property Get vbdObject_AlingText() As Integer
        vbdObject_AlingText = m_Object.AlingText
End Property

Private Property Let vbdObject_AlingText(ByVal RHS As Integer)
       m_Object.AlingText = RHS
End Property

Private Property Let vbdObject_Angle(ByVal RHS As Single)
        m_Object.Angle = RHS
End Property

Private Property Get vbdObject_Angle() As Single
          vbdObject_Angle = m_Object.Angle
End Property

Public Property Let vbdObject_Bold(ByVal RHS As Boolean)
       m_Object.Bold = RHS
End Property

Public Property Get vbdObject_Bold() As Boolean
       vbdObject_Bold = m_Object.Bold
End Property

Private Property Set vbdObject_Canvas(ByVal RHS As PictureBox)
    Set m_Canvas = RHS
End Property

Private Property Get vbdObject_Canvas() As PictureBox
      Set vbdObject_Canvas = m_Canvas
End Property

Public Property Let vbdObject_Charset(ByVal RHS As Integer)
         m_Object.Charset = RHS
End Property

Public Property Get vbdObject_Charset() As Integer
         vbdObject_Charset = m_Object.Charset
End Property

' Clear the object's transformation.
Private Sub vbdObject_ClearTransformation()
     m_Object.ClearTransformation
End Sub

' Add this transformation to the current one.
Private Sub vbdObject_AddTransformation(m() As Single)
    m_Object.AddTransformation m
    m_Object.MakeTransformation
End Sub
Public Property Let vbdObject_CurrentX(ByVal RHS As Single)
      m_Object.CurrentX = RHS
End Property

Public Property Get vbdObject_CurrentX() As Single
       vbdObject_CurrentX = m_Object.CurrentX
End Property

Public Property Let vbdObject_CurrentY(ByVal RHS As Single)
      m_Object.CurrentY = RHS
End Property

Public Property Get vbdObject_CurrentY() As Single
         vbdObject_CurrentY = m_Object.CurrentY
End Property

' Draw the object in a metafile.
Private Sub vbdObject_DrawInMetafile(ByVal mf_dc As Long)
     m_Object.DrawInMetafile mf_dc
End Sub

Private Sub vbdObject_DrawPoint()
    m_DrawingStarted = False
    m_DrawingEdit = True
    m_Canvas.MousePointer = 2
    DrawPoint True
End Sub

' Draw box on the object.
Private Sub vbdObject_DrawSelect(ByVal pic As PictureBox)
        m_Object.DrawSelect pic
End Sub

' Return the object's DrawWidth.
Public Property Get vbdObject_DrawWidth() As Integer
    vbdObject_DrawWidth = m_Object.DrawWidth
End Property
' Set the object's DrawWidth.
Public Property Let vbdObject_DrawWidth(ByVal new_value As Integer)
    m_Object.DrawWidth = new_value
End Property

' Return the object's DrawStyle.
Public Property Get vbdObject_DrawStyle() As Integer
    vbdObject_DrawStyle = m_Object.DrawStyle
End Property
' Set the object's DrawStyle.
Public Property Let vbdObject_DrawStyle(ByVal new_value As Integer)
    m_Object.DrawStyle = new_value
End Property

Private Property Let vbdObject_EditPoint(ByVal RHS As Boolean)
         m_Object.EditPoint = RHS
End Property

Private Property Get vbdObject_EditPoint() As Boolean
       vbdObject_EditPoint = m_Object.EditPoint
End Property

Public Property Let vbdObject_FillColor2(ByVal RHS As stdole.OLE_COLOR)
     m_Object.FillColor2 = RHS
End Property

Public Property Get vbdObject_FillColor2() As stdole.OLE_COLOR
    vbdObject_FillColor2 = m_Object.FillColor2
End Property

Public Property Let vbdObject_FillMode(ByVal RHS As FillMode)
     m_Object.FillMode = RHS
End Property

Public Property Get vbdObject_FillMode() As FillMode
        vbdObject_FillMode = m_Object.FillMode
End Property

' Return the object's ForeColor.
Public Property Get vbdObject_ForeColor() As OLE_COLOR
    vbdObject_ForeColor = m_Object.ForeColor
End Property
' Set the object's ForeColor.
Public Property Let vbdObject_ForeColor(ByVal new_value As OLE_COLOR)
    m_Object.ForeColor = new_value
End Property

' Return the object's FillColor.
Public Property Get vbdObject_FillColor() As OLE_COLOR
    vbdObject_FillColor = m_Object.FillColor
End Property
' Set the object's FillColor.
Public Property Let vbdObject_FillColor(ByVal new_value As OLE_COLOR)
    m_Object.FillColor = new_value
End Property

' Return the object's FillStyle.
Public Property Get vbdObject_FillStyle() As Integer
    vbdObject_FillStyle = m_Object.FillStyle
End Property
' Set the object's FillStyle.
Public Property Let vbdObject_FillStyle(ByVal new_value As Integer)
    m_Object.FillStyle = new_value
End Property

' Return this object's bounds.
Public Sub vbdObject_Bound(ByRef xmin As Single, ByRef ymin As Single, ByRef xmax As Single, ByRef ymax As Single)
       m_Object.Bound xmin, ymin, xmax, ymax
End Sub
' Draw the object on the canvas.
Public Sub vbdObject_Draw(ByVal pic As PictureBox, Optional SetWsc As Boolean = True)
    m_Object.Draw pic
End Sub

Public Property Let vbdObject_Gradient(ByVal RHS As Integer)
      m_Object.Gradient = RHS
End Property

Public Property Get vbdObject_Gradient() As Integer
         vbdObject_Gradient = m_Object.Gradient
End Property

Private Property Let vbdObject_hRegion(ByVal RHS As Long)
          m_Object.hRegion = RHS
End Property

Private Property Get vbdObject_hRegion() As Long
         vbdObject_hRegion = m_Object.hRegion
End Property

Public Function vbdObject_Info() As String
       vbdObject_Info = m_Object.Info
End Function

Public Property Get vbdObject_Italic() As Boolean
         vbdObject_Italic = m_Object.Italic
End Property
Public Property Let vbdObject_Italic(ByVal RHS As Boolean)
           m_Object.Italic = RHS
End Property

Public Sub vbdObject_MakeTransformation()
      m_Object.MakeTransformation
End Sub

Public Property Get vbdObject_Name() As String
         vbdObject_Name = m_Object.Name
End Property

Public Property Let vbdObject_Name(ByVal RHS As String)
        m_Object.Name = RHS
End Property

Private Sub vbdObject_NewPoint(NumPoints As Long, X() As Single, Y() As Single, TypePoint() As Byte)
          m_Object.NewPoint NumPoints, X(), Y(), TypePoint()

'      Dim n As Long
'      m_Polygon.NumPoints = NumPoints
'
'      For n = 1 To NumPoints
'          m_Polygon.X(n) = X(n) '+ m_Object.CurrentX
'          m_Polygon.Y(n) = Y(n) '+ m_Object.CurrentY
'          m_Polygon.TypePoint(n) = TypePoint(n)
'      Next
'       m_Polygon.Identity
End Sub

Private Sub vbdObject_NewTrPoint(NumPoints As Long, X() As Single, Y() As Single, TypePoint() As Byte)
      Dim n As Long
      m_NumPoints = NumPoints
      ReDim m_OriginalPoints(1 To NumPoints)
      ReDim m_TypePoint(1 To NumPoints)
      For n = 1 To m_NumPoints
          m_OriginalPoints(n).X = X(n)
          m_OriginalPoints(n).Y = Y(n)
          m_TypePoint(n) = TypePoint(n)
      Next
       m_Polygon.Identity
End Sub

Public Property Let vbdObject_ObjLock(ByVal RHS As Boolean)
         m_Object.ObjLock = RHS
End Property

Public Property Get vbdObject_ObjLock() As Boolean
         vbdObject_ObjLock = m_Object.ObjLock
End Property

Public Property Let vbdObject_Blend(ByVal RHS As Integer)
           m_Object.Blend = RHS
End Property

Public Property Get vbdObject_Blend() As Integer
         vbdObject_Blend = m_Object.Blend
End Property

Public Property Let vbdObject_Pattern(ByVal RHS As String)
       m_Object.Pattern = RHS
End Property

Public Property Get vbdObject_Pattern() As String
       vbdObject_Pattern = m_Object.Pattern
End Property

Public Property Get vbdObject_Picture() As stdole.Picture
       Set vbdObject_Picture = m_Object.Picture
End Property
Public Property Set vbdObject_Picture(ByVal RHS As stdole.Picture)
       Set m_Object.Picture = RHS
End Property

Private Sub vbdObject_ReadPoint(ByRef NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
       m_Object.ReadPoint NumPoints, X(), Y(), TPoint()
End Sub

Private Sub vbdObject_ReadTrPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
       m_Object.ReadTrPoint NumPoints, X(), Y(), TPoint()
End Sub

' Set the object's Selected status.
Public Property Let vbdObject_Selected(ByVal RHS As Boolean)
    m_Object.Selected = RHS
End Property
' Return the object's Selected status.
Public Property Get vbdObject_Selected() As Boolean
    vbdObject_Selected = m_Object.Selected
End Property

' Return True if the object is at this location.
Private Function vbdObject_IsAt(ByVal X As Single, ByVal Y As Single) As Boolean
     vbdObject_IsAt = m_Object.IsAt(X, Y)
End Function

' Initialize the object using a serialization string.
' The serialization does not include the
' ObjectType(...) part.
Private Property Let vbdObject_Serialization(ByVal RHS As String)
Dim token_name As String
Dim token_value As String
Dim next_x As Integer
Dim next_y As Integer

    ' Start with a new polygon.
    Set m_Polygon = New vbdDraw
    Set m_Object = m_Polygon

    ' Read tokens until there are no more.
    Do While Len(RHS) > 0
        ' Read a token.
        GetNamedToken RHS, token_name, token_value
        Select Case token_name
            Case "IsClosed"
                IsClosed = CBool(token_value)
            Case "IsBox"
                IsBox = CBool(token_value)
            Case "Polyline", "Scribble", "FreePolygon", "Polygon", "Curve", "PolyDraw", "Picture", "RectAngle", "Ellipse", "Text", "TextFrame", "TextPath", "Scene"
                If token_name = "PolyDraw" Or token_name = "Curve" Then m_Object.TypeDraw = dPolydraw: Me.m_TypeDraw = dPolydraw
                m_Object.Serialization = token_value
        End Select
    Loop
End Property

' Return a serialization string for the object.
Public Property Get vbdObject_Serialization() As String
Dim txt As String
    If Me.m_TypeDraw = dFreePolygon Or Me.m_TypeDraw = dPolyline Then
       txt = txt & " IsClosed(" & Format$(IsClosed) & ")"
    End If
    txt = txt & m_Object.Serialization

    ''Debug.Print m_Object.TypeDraw = dFreePolygon
    Select Case Me.m_TypeDraw
    Case dFreePolygon
        vbdObject_Serialization = "FreePolygon(" & txt & ")"
    Case dPolyline
        vbdObject_Serialization = "Polyline(" & txt & ")"
    Case dPolydraw
        vbdObject_Serialization = txt
    Case dScribble
    End Select

End Property

Public Property Let vbdObject_Shade(ByVal RHS As Boolean)
       m_Object.Shade = RHS
End Property

Public Property Get vbdObject_Shade() As Boolean
      vbdObject_Shade = m_Object.Shade
End Property

Public Property Let vbdObject_Size(ByVal RHS As Integer)
        m_Object.Size = RHS
End Property

Public Property Get vbdObject_Size() As Integer
          vbdObject_Size = m_Object.Size
End Property

Public Sub vbdObject_StartBound(xmin As Single, ymin As Single)
     m_Object.StartBound xmin, ymin
End Sub

Public Property Let vbdObject_Strikethrough(ByVal RHS As Boolean)
        m_Object.Strikethrough = RHS
End Property
Public Property Get vbdObject_Strikethrough() As Boolean
         vbdObject_Strikethrough = m_Object.Strikethrough
End Property

Public Property Get vbdObject_TextDraw() As String
        vbdObject_TextDraw = m_Object.TextDraw
End Property

Public Property Let vbdObject_TextDraw(ByVal RHS As String)
     m_Object.TextDraw = RHS
End Property

' Set the object's TypeDraw.
Public Property Let vbdObject_TypeDraw(ByVal RHS As DrawType)
     m_Object.TypeDraw = RHS
End Property
' Return the object's TypeDraw.
Public Property Get vbdObject_TypeDraw() As DrawType
       vbdObject_TypeDraw = m_Object.TypeDraw
End Property

Public Property Let vbdObject_TypeFill(ByVal RHS As DrawTypeFill)
       m_Object.TypeFill = RHS
End Property

Public Property Get vbdObject_TypeFill() As DrawTypeFill
       vbdObject_TypeFill = m_Object.TypeFill
End Property

Public Property Let vbdObject_Underline(ByVal RHS As Boolean)
     m_Object.Underline = RHS
End Property

Public Property Get vbdObject_Underline() As Boolean
      vbdObject_Underline = m_Object.Underline
End Property

Public Property Let vbdObject_Weight(ByVal RHS As Integer)
        m_Object.Weight = RHS
End Property

Public Property Get vbdObject_Weight() As Integer
        vbdObject_Weight = m_Object.Weight
End Property
