Attribute VB_Name = "Module1"

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
  Data4(7) As Byte
End Type
Private Type PicDesc
  Size As Long
  Type As Long
  hBmp As Long
  hPal As Long
  Reserved As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (pDesc As PicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, pPic As IPicture) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub HiMetricToPixel Lib "atl" Alias "AtlHiMetricToPixel" (lpSizeInHiMetric As SIZEL, lpSizeInPix As SIZEL)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Const ICON_OFFSET As Long = 32
Private Type SIZEL
     cx As Long
     cy As Long
End Type
Public Sub DrawPicture(ByVal plnghDC As Long, ByVal plngPicture As Long, ByVal plngX As Long, ByVal plngY As Long, ByVal plngWidth As Long, ByVal plngHeight As Long)
Dim tHIMETRIC   As SIZEL
Dim tPIXEL      As SIZEL
Dim lngDC       As Long
Dim oPicture    As StdPicture
Dim lSrcDC As Long
Dim lBMP As Long
Dim lOldObj As Long
Dim lOldSrcObj As Long
Dim lngXPos     As Long
Dim lngYpos     As Long
Dim lngWidth    As Long
Dim lngHeight   As Long
    Set oPicture = CreateBitmapPicture(plngPicture)    ' Create a compatible DC
    lngDC = CreateCompatibleDC(plnghDC)      ' Create a compatible Bitmap
    lBMP = CreateCompatibleBitmap(plnghDC, 3, 3) 'plngWidth, plngHeight)      ' Select the new Bitmap into the new DC
    lOldObj = SelectObject(lngDC, lBMP)        ' Create a Compatible DC for the Source Image
    lSrcDC = CreateCompatibleDC(plnghDC)
    SetBkMode lSrcDC, TRANSPARENT      ' Select the Loaded Bitmap into the DC
    lOldSrcObj = SelectObject(lSrcDC, oPicture.Handle)     ' Convert the HiMetric dimensions of the Image to Pixels
    tHIMETRIC.cx = oPicture.Width
    tHIMETRIC.cy = oPicture.Height
    Call HiMetricToPixel(tHIMETRIC, tPIXEL)
        For lngYpos = plngY To plngY + plngHeight Step tPIXEL.cy
        For lngXPos = plngX To plngX + plngWidth Step tPIXEL.cx
            If (plngX + plngWidth) - lngXPos < tPIXEL.cx Then
                lngWidth = (plngX + plngWidth) - lngXPos
            Else
                lngWidth = tPIXEL.cx
            End If
            If (plngY + plngHeight) - lngYpos < tPIXEL.cy Then
                lngHeight = (plngY + plngHeight) - lngYpos
            Else
                lngHeight = tPIXEL.cy
            End If
            BitBlt plnghDC, lngXPos, lngYpos, lngWidth, lngHeight, lSrcDC, 0, 0, vbSrcAnd
            BitBlt plnghDC, lngXPos, lngYpos, lngWidth, lngHeight, lSrcDC, 0, 0, vbSrcPaint
         Next lngXPos
        Next lngYpos   ' Clean up the Source Image DC and Bitmap
  Call SelectObject(lSrcDC, lOldSrcObj)
  Call DeleteDC(lSrcDC)    ' Select the New bitmap out of the new DC
  lBMP = SelectObject(lngDC, lOldObj)    ' Delete the DC
  Call DeleteDC(lngDC)
End Sub

Private Function CreateBitmapPicture(ByVal lBMP As Long) As Picture
  Dim tPic As PicDesc
  Dim oPic As IPicture
  Dim IID_IDispatch As GUID
    With IID_IDispatch
     .Data1 = &H20400
     .Data4(0) = &HC0
     .Data4(7) = &H46
    End With
    With tPic
     .Size = Len(tPic)
     .Type = 1
     .hBmp = lBMP
     .hPal = 0
    End With    ' Create Picture object.
  OleCreatePictureIndirect tPic, IID_IDispatch, False, oPic    ' Return the new Picture object.
  Set CreateBitmapPicture = oPic
End Function



'Private Sub PolyPoints(nPoint As Integer, cx As Single, cy As Single)
'    If nPoint > 0 Then
'        m_OriginalPoints(nPoint).X = cx
'        m_OriginalPoints(nPoint).Y = cy
'    End If
'End Sub
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
'' Calculate new Point in the line.
'Public Function AddNode(ByVal X As Single, ByVal Y As Single) As POINTAPI
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
'               If MinX + bStep >= m_OriginalPoints(I).X And MinX - bStep <= m_OriginalPoints(I).X Then
'               If MinY + bStep >= m_OriginalPoints(I).Y And MinY - bStep <= m_OriginalPoints(I).Y Then
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
'Public Sub DeleteNode()
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
'Public Sub toLine()
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
'Public Sub toCurve()
'     Dim Points() As POINTAPI, aa As Long, mType As Byte, I As Long
'     Dim mTypePoint() As Byte, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single
'     If m_NumPoints > 2 Then
'
'         If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
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
'              Points(m_SelectPoint).X = X3
'              Points(m_SelectPoint).Y = Y3
'              Points(m_SelectPoint + 1).X = X2
'              Points(m_SelectPoint + 1).Y = Y2
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
'             m_OriginalPoints = Points
'             m_TypePoint = mTypePoint
'         End If
'     End If
'End Sub
'
''Close Node line
'Public Sub CloseNode()
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
'Public Sub OpenNode()
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
'Public Sub BreakNode()
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
'''Popup Menu
''Public Function MenuNode() As Long
''    Dim Pt As POINTAPI
''    Dim ret As Long
''    Dim wFlag0 As Long, wFlag1 As Long, wFlag2 As Long, wFlag3 As Long, wFlag4 As Long, wFlag5 As Long, wFlag6 As Long
''
''    If IsControl(m_TypePoint, m_SelectPoint) = True Then
''       wFlag0 = MF_GRAYED Or MF_DISABLED
''    Else
''        wFlag0 = MF_STRING
''    End If
''    If m_SelectPoint = 0 Or m_SelectPoint = 1 Or IsControl(m_TypePoint, m_SelectPoint) = True Then
''        wFlag1 = MF_GRAYED Or MF_DISABLED
''    Else
''        wFlag1 = MF_STRING
''    End If
''
''    If m_SelectPoint > 0 Then
''       If m_TypePoint(m_SelectPoint) = 6 Then
''          If m_SelectPoint + 1 > m_NumPoints Then
''               wFlag2 = MF_GRAYED Or MF_DISABLED
''              wFlag3 = MF_GRAYED Or MF_DISABLED
''          Else
''          If m_TypePoint(m_SelectPoint + 1) = 2 Then wFlag2 = MF_GRAYED Or MF_DISABLED Else wFlag2 = MF_STRING
''          If m_TypePoint(m_SelectPoint + 1) = 4 Then wFlag3 = MF_GRAYED Or MF_DISABLED Else wFlag3 = MF_STRING
''          End If
''       ElseIf m_TypePoint(m_SelectPoint) = 3 Then
''           wFlag2 = MF_GRAYED Or MF_DISABLED
''           wFlag3 = MF_GRAYED Or MF_DISABLED
''       Else
''          If m_SelectPoint >= m_NumPoints Then m_SelectPoint = m_SelectPoint - 1
''          If m_TypePoint(m_SelectPoint + 1) = 2 Or _
''             m_TypePoint(m_SelectPoint + 1) = 6 Or _
''             m_TypePoint(m_SelectPoint + 1) = 3 Or _
''             IsControl(m_TypePoint, m_SelectPoint) = True Then
''              wFlag2 = MF_GRAYED Or MF_DISABLED
''          Else
''              wFlag2 = MF_STRING
''          End If
''
''          If m_TypePoint(m_SelectPoint + 1) = 4 Or _
''             m_TypePoint(m_SelectPoint + 1) = 6 Or _
''             m_TypePoint(m_SelectPoint + 1) = 3 Or _
''             IsControl(m_TypePoint, m_SelectPoint) = True Then
''             wFlag3 = MF_GRAYED Or MF_DISABLED
''          Else
''             wFlag3 = MF_STRING
''          End If
''
''       End If
''    Else
''       wFlag2 = MF_GRAYED Or MF_DISABLED
''       wFlag3 = MF_GRAYED Or MF_DISABLED
''    End If
''
''    If m_TypePoint(m_NumPoints) = 3 Or m_TypePoint(m_NumPoints) = 5 Then wFlag4 = MF_GRAYED Or MF_DISABLED Else wFlag4 = MF_STRING
''
''    If IsOpening(m_TypePoint) Then wFlag5 = MF_STRING Else wFlag5 = MF_GRAYED Or MF_DISABLED
''
''    If m_SelectPoint > 0 And m_SelectPoint <> m_NumPoints And m_TypePoint(m_NumPoints) <> 6 And _
''       IsControl(m_TypePoint, m_SelectPoint) = False Then
''       wFlag6 = MF_STRING
''    Else
''       wFlag6 = MF_GRAYED Or MF_DISABLED
''    End If
''
''    hMenu = CreatePopupMenu()
''    AppendMenu hMenu, wFlag0, 1, "Add node(s)"
''    AppendMenu hMenu, wFlag1, 2, "Delete node(s)" + vbTab + "(Del)"
''    AppendMenu hMenu, MF_SEPARATOR, 3, ByVal 0&
''    AppendMenu hMenu, wFlag2, 4, "To Line"
''    AppendMenu hMenu, wFlag3, 5, "To Curve"
''    AppendMenu hMenu, MF_SEPARATOR, 6, ByVal 0&
''    AppendMenu hMenu, wFlag4, 7, "Auto Close"
''    AppendMenu hMenu, wFlag5, 8, "Auto Open"
''    AppendMenu hMenu, MF_SEPARATOR, 9, ByVal 0&
''    AppendMenu hMenu, wFlag6, 10, "Break node"
''    AppendMenu hMenu, MF_GRAYED Or MF_DISABLED, 11, "Break Apart"
''
''    ''Debug.Print m_TypePoint(m_NumPoints)
''    GetCursorPos Pt
''    ret = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, Pt.X, Pt.Y, m_Canvas.hWnd, ByVal 0&)
''    DestroyMenu hMenu
''   ''Debug.Print ret
''    MenuNode = ret
''End Function
'
'
''Check select point if is Control
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
'' Check is opening the line
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
''
'Public Sub FindNodeCurve(ByVal px1 As Single, ByVal py1 As Single, ByVal px2 As Single, ByVal py2 As Single, _
'                          ByRef cX1 As Single, ByRef cY1 As Single, ByRef cX2 As Single, ByRef cY2 As Single)
'            Dim tX1 As Single, tY1 As Single
'
'            MidPoint px1, py1, px2, py2, tX1, tY1
'            MidPoint px1, py1, tX1, tY1, cX1, cY1
'            MidPoint tX1, tY1, px2, py2, cX2, cY2
'End Sub
'

