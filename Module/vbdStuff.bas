Attribute VB_Name = "vbdStuff"
Option Explicit

Private m_OldPen As Long
Private m_OldBrush As Long
Private m_NewBrush As Long
Private m_NewPen As Long

' Bound the objects in the collection.
Public Sub BoundObjects(ByVal the_objects As Collection, ByRef xmin As Single, ByRef ymin As Single, ByRef xmax As Single, ByRef ymax As Single)
Dim X1 As Single
Dim X2 As Single
Dim Y1 As Single
Dim Y2 As Single
Dim Obj As vbdObject
Dim r As RECT
    'If the_objects(1) Is Nothing Then Exit Sub
    Set Obj = the_objects(1)
   
    GetRgnBox Obj.hRegion, r
    xmin = r.Left
    ymin = r.Top
    xmax = r.Right
    ymax = r.Bottom
    If Obj.TypeDraw = dPolydraw Then
       Obj.Bound xmin, ymin, xmax, ymax
    End If
End Sub

Public Sub NewTransformation()
    Dim the_scene As vbdScene
    ' Save the new object.
    Set the_scene = m_TheScene
    the_scene.NewTransformation
    Set the_scene = Nothing
End Sub

' Return this object's bounds.
Public Sub BoundText(ByRef Points() As PointAPI, ByRef xmin As Single, ByRef ymin As Single, ByRef xmax As Single, ByRef ymax As Single)
Dim i As Integer, m_NumPoints As Long
    m_NumPoints = UBound(Points)
    If m_NumPoints < 1 Then
        xmin = 0
        xmax = 0
        ymin = 0
        ymax = 0
    Else
        With Points(1)
            xmin = .X
            xmax = xmin
            ymin = .Y
            ymax = ymin
        End With

        For i = 2 To m_NumPoints
            With Points(i)
                If xmin > .X Then xmin = .X
                If xmax < .X Then xmax = .X
                If ymin > .Y Then ymin = .Y
                If ymax < .Y Then ymax = .Y
            End With
        Next i
    End If
End Sub

' StartBound the objects in the collection.
Public Sub StartBoundObjects(ByVal the_objects As Collection, ByRef xmin As Single, ByRef ymin As Single)
    Dim Obj As vbdObject
    Set Obj = the_objects(1)
    Obj.StartBound xmin, ymin
End Sub

' Initialize default drawing properties.
Public Sub InitializeDrawingProperties(ByVal Obj As vbdObject)
    Obj.DrawWidth = 1
    Obj.DrawStyle = vbSolid
    Obj.ForeColor = vbBlack
    Obj.FillColor = vbBlack
    Obj.FillStyle = vbFSTransparent
    Obj.FillMode = fALTERNATE
    Obj.TextDraw = ""
   ' Obj.TypeDraw = 0
    Obj.TypeFill = 0
    Obj.Bold = False
    Obj.Charset = 0
    Obj.Italic = False
    Obj.Name = "Arial"
    Obj.Size = 20
    Obj.Strikethrough = False
    Obj.Underline = False
    Obj.Weight = 400
    Obj.Angle = 0
    Obj.Pattern = ""
End Sub
' Return the drawing property serialization
' for this object.
Public Function DrawingPropertySerialization(ByVal Obj As vbdObject) As String
Dim txt As String

    txt = txt & " DrawWidth(" & Format$(Obj.DrawWidth) & ")"
    txt = txt & " DrawStyle(" & Format$(Obj.DrawStyle) & ")"
    txt = txt & " ForeColor(" & Format$(Obj.ForeColor) & ")"
    txt = txt & " FillColor(" & Format$(Obj.FillColor) & ")"
    txt = txt & " FillColor2(" & Format$(Obj.FillColor2) & ")"
    txt = txt & " FillMode(" & Format$(Obj.FillMode) & ")"
    txt = txt & " Pattern(" & Trim(Obj.Pattern) & ")"
    txt = txt & " Gradient(" & Format$(Obj.Gradient) & ")"
    txt = txt & " FillStyle(" & Format$(Obj.FillStyle) & ")"
    txt = txt & " TextDraw(" & Format$(Obj.TextDraw) & ")"
    txt = txt & " TypeDraw(" & Format$(Obj.TypeDraw) & ")"
    txt = txt & " CurrentX(" & Format$(Obj.CurrentX) & ")"
    txt = txt & " CurrentY(" & Format$(Obj.CurrentY) & ")"
    txt = txt & " TypeFill(" & Format$(Obj.TypeFill) & ")"
    txt = txt & " Shade(" & Format$(Obj.Shade) & ")"
    txt = txt & " ObjLock(" & Format$(Obj.ObjLock) & ")"
    txt = txt & " Blend(" & Format$(Obj.Blend) & ")"
    
    txt = txt & " Bold(" & Format$(Obj.Bold) & ")"
    txt = txt & " Charset(" & Format$(Obj.Charset) & ")"
    txt = txt & " Italic(" & Format$(Obj.Italic) & ")"
    txt = txt & " Name(" & Format$(Obj.Name) & ")"
    txt = txt & " Size(" & Format$(Obj.Size) & ")"
    txt = txt & " Strikethrough(" & Format$(Obj.Strikethrough) & ")"
    txt = txt & " Underline(" & Format$(Obj.Underline) & ")"
    txt = txt & " Weight(" & Format$(Obj.Weight) & ")"
    txt = txt & " Angle(" & Format$(Obj.Angle) & ")"
    
    DrawingPropertySerialization = txt & vbCrLf & "    "
End Function

' Read the token name and value and to see
' if it is drawing property information.
Public Sub ReadDrawingPropertySerialization(ByVal Obj As vbdObject, ByVal token_name As String, ByVal token_value As String)
    
    Select Case token_name
        Case "DrawWidth"
            Obj.DrawWidth = CInt(token_value)
        Case "DrawStyle"
            Obj.DrawStyle = CInt(token_value)
        Case "ForeColor"
            Obj.ForeColor = CLng(token_value)
        Case "FillColor"
            Obj.FillColor = CLng(token_value)
        Case "FillColor2"
           Obj.FillColor2 = CLng(token_value)
         Case "FillMode"
            Obj.FillMode = CSng(token_value)
        Case "Pattern"
           Obj.Pattern = Trim(token_value)
        Case "Gradient"
           Obj.Gradient = CInt(token_value)
        Case "FillStyle"
            Obj.FillStyle = CInt(token_value)
        Case "TextDraw"
            Obj.TextDraw = Trim(token_value)
        Case "TypeDraw"
            Obj.TypeDraw = CInt(token_value)
        Case "TypeFill"
            Obj.TypeFill = CInt(token_value)
        Case "Shade"
            Obj.Shade = Val(token_value)
        Case "ObjLock"
            Obj.ObjLock = CBool(token_value)
        Case "Blend", "Opacity"
            Obj.Blend = Val(token_value)
        Case "Angle"
            Obj.Angle = CSng(token_value)
        Case "Charset"
            Obj.Charset = CInt(token_value)
        Case "Italic"
            Obj.Italic = CBool(token_value)
        Case "Name"
            Obj.Name = Trim(token_value)
        Case "Size"
            Obj.Size = CInt(token_value)
        Case "Bold"
            Obj.Bold = CBool(token_value)
        Case "Strikethrough"
            Obj.Strikethrough = CBool(token_value)
        Case "Underline"
            Obj.Underline = CBool(token_value)
        Case "Weight"
            Obj.Weight = CLng(token_value)
        Case "CurrentX"
            Obj.CurrentX = CSng(token_value)
        Case "CurrentY"
            Obj.CurrentY = CSng(token_value)
        Case "AlingText"
            Obj.AlingText = CSng(token_value)
        Case Else
         ' Stop
    End Select
End Sub


' Set the drawing properties for the metafile.
Public Sub SetMetafileDrawingParameters(ByVal Obj As vbdObject, ByVal mf_dc As Long)
Dim log_brush As LogBrush
Dim new_brush As Long
Dim new_pen As Long

    With log_brush
        If Obj.FillStyle = vbFSTransparent Then
            .lbStyle = BS_HOLLOW
        ElseIf Obj.FillStyle = vbFSSolid Then
            .lbStyle = BS_SOLID
        Else
            .lbStyle = BS_HATCHED
            Select Case Obj.FillStyle
                Case vbCross
                    .lbHatch = HS_CROSS
                Case vbDiagonalCross
                    .lbHatch = HS_DIAGCROSS
                Case vbDownwardDiagonal
                    .lbHatch = HS_BDIAGONAL
                Case vbHorizontalLine
                    .lbHatch = HS_HORIZONTAL
                Case vbUpwardDiagonal
                    .lbHatch = HS_FDIAGONAL
                Case vbVerticalLine
                    .lbHatch = HS_VERTICAL
            End Select
        End If
        .lbColor = Obj.FillColor
    End With

    m_NewPen = CreatePen(Obj.DrawStyle, Obj.DrawWidth, Obj.ForeColor)
    m_NewBrush = CreateBrushIndirect(log_brush)
    m_OldPen = SelectObject(mf_dc, m_NewPen)
    m_OldBrush = SelectObject(mf_dc, m_NewBrush)
End Sub

' Restore the drawing properties for the metafile.
Public Sub RestoreMetafileDrawingParameters(ByVal mf_dc As Long)
    SelectObject mf_dc, m_OldBrush
    SelectObject mf_dc, m_OldPen
    DeleteObject m_NewBrush
    DeleteObject m_NewPen
End Sub

' Return the serialization for this transformation matrix.
Public Function TransformationSerialization(m() As Single) As String
Dim i As Integer
Dim j As Integer
Dim txt As String

    For i = 1 To 3
        For j = 1 To 3
            txt = txt & Format$(m(i, j)) & " "
        Next j
    Next i

    TransformationSerialization = "Transformation(" & txt & ")"
End Function

' initialize the transformation matrix using this serialization.
Public Sub SetTransformationSerialization(ByVal txt As String, m() As Single)
Dim i As Integer
Dim j As Integer
Dim Token As String

    For i = 1 To 3
        For j = 1 To 3
            Token = GetDelimitedToken(txt, " ")
            Token = Replace(Token, ",", ".")
            m(i, j) = CSng(Val(Token))
        Next j
    Next i
End Sub

Public Function LoadPatternPic(FileImage As String, _
                              Optional dWidth As Long = 8, _
                              Optional dHeight As Long = 8) As Long
     Const LR_LOADFROMFILE = &H10
     Const IMAGE_BITMAP = 0

     If FileExists(FileImage) Then
        LoadPatternPic = LoadImage(App.hInstance, FileImage, IMAGE_BITMAP, dWidth, dHeight, LR_LOADFROMFILE)
     End If
End Function

'Create pen Style
Public Function PenCreate(mDrawStyle As Integer, mWidthLine As Integer, mColorLine As Long) As Long
Dim BrushInf As LogBrush
Dim StyleArr() As Long
Dim wLine As Long
Dim PenStyle As Long
    
    wLine = mWidthLine
        
    Select Case mDrawStyle
    Case 0 'vbSolid
       ReDim StyleArr(1)
       StyleArr(0) = 10 '* (wLine / 2)
       StyleArr(1) = 0 '* (wLine / 2)
        'PenCreate = CreatePen(PS_SOLID, wLine, mColorLine)
        'Exit Function
    Case 1 'vbDash
       ReDim StyleArr(1)
        StyleArr(0) = 18 * (wLine / 2)
        StyleArr(1) = 6 * (wLine / 2)

    Case 2 'vbDot
       ReDim StyleArr(3)
        StyleArr(0) = 3 * (wLine / 2)
        StyleArr(1) = 3 * (wLine / 2)
        StyleArr(2) = 3 * (wLine / 2)
        StyleArr(3) = 3 * (wLine / 2)
        
    Case 3 'vbDashDot
       ReDim StyleArr(3)
        StyleArr(0) = 9 * (wLine / 2)
        StyleArr(1) = 6 * (wLine / 2)
        StyleArr(2) = 3 * (wLine / 2)
        StyleArr(3) = 6 * (wLine / 2)
    
    Case 4 'vbDashDotDot
        ReDim StyleArr(5)
        StyleArr(0) = 9 * (wLine / 2)
        StyleArr(1) = 3 * (wLine / 2)
        StyleArr(2) = 3 * (wLine / 2)
        StyleArr(3) = 3 * (wLine / 2)
        StyleArr(4) = 3 * (wLine / 2)
        StyleArr(5) = 3 * (wLine / 2)
        
    Case 5 'vbInvisible
        PenCreate = CreatePen(PS_NULL, wLine, mColorLine)
        Exit Function
    End Select
    
    BrushInf.lbColor = mColorLine
    PenCreate = ExtCreatePen(PS_GEOMETRIC Or PS_USERSTYLE Or PS_ENDCAP_ROUND, wLine, BrushInf, UBound(StyleArr()) + 1, StyleArr(0))
    
    Erase StyleArr
    
End Function

Public Function BitmapFromDC(ByVal lhDC As Long, _
                             ByVal lLeft As Long, _
                             ByVal lTop As Long, _
                             ByVal lWidth As Long, _
                             ByVal lHeight As Long) As Long

   ' Copy the bitmap in lHDC:
   Dim lhDCCopy As Long
   Dim lhBmpCopy As Long
   Dim lhBmpCopyOld As Long
   Dim lhDCC As Long
   Dim tBM As BITMAP
   
   lhDCC = CreateDCA("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   lhDCCopy = CreateCompatibleDC(lhDCC)
   
   lhBmpCopy = CreateCompatibleBitmap(lhDCC, lWidth, lHeight)
   lhBmpCopyOld = SelectObject(lhDCCopy, lhBmpCopy)
   Call BitBlt(lhDCCopy, lLeft, lTop, lWidth, lHeight, lhDC, 0, 0, vbSrcCopy)
   
   If Not (lhDCC = 0) Then
      DeleteDC lhDCC
   End If
   If Not (lhBmpCopyOld = 0) Then
      SelectObject lhDCCopy, lhBmpCopyOld
   End If
   If Not (lhDCCopy = 0) Then
      DeleteDC lhDCCopy
   End If

   BitmapFromDC = lhBmpCopy

End Function

' Return the next delimited token from txt. Trim blanks.
Public Function GetDelimitedToken(ByRef txt As String, ByVal delimiter As String) As String
Dim pos As Integer

    pos = InStr(txt, delimiter)
    If pos < 1 Then
        ' The delimiter was not found. Return the rest of txt.
        GetDelimitedToken = Trim$(txt)
        txt = ""
    Else
        ' We found the delimiter. Return the token.
        GetDelimitedToken = Trim$(Left$(txt, pos - 1))
        txt = Trim$(mid$(txt, pos + Len(delimiter)))
    End If
End Function

' Replace non-printable characters in txt with spaces.
Public Function NonPrintingToSpace(ByVal txt As String) As String
Dim i As Integer
Dim cH As String

    For i = 1 To Len(txt)
        cH = mid$(txt, i, 1)
        If (cH < " ") Or (cH > "~") Then Mid$(txt, i, 1) = " "
    Next i
    NonPrintingToSpace = txt
End Function


' Remove comments starting with  from the end of lines.
Public Function RemoveComments(ByVal txt As String) As String
Dim pos As Integer
Dim new_txt As String

    Do While Len(txt) > 0
        ' Find the next '.
        pos = InStr(txt, "'")
        If pos = 0 Then
            new_txt = new_txt & txt
            Exit Do
        End If

        ' Add this part to the result.
        new_txt = new_txt & Left$(txt, pos - 1)

        ' Find the end of the line.
        pos = InStr(pos + 1, txt, vbCrLf)
        If pos = 0 Then
            ' There was no vbCrLf. Remove the rest of the text.
            txt = ""
        Else
            txt = mid$(txt, pos + Len(vbCrLf))
        End If
    Loop

    RemoveComments = new_txt
End Function


'Public Function PolygonPoints(cPtsQty As Integer, cLeft As Single, cTop As Single, cWidth As Single, cHeight As Single) As POINTAPI()
'
'Dim POINT() As POINTAPI
'Dim n As Integer
'Dim RadiusW As Single
'Dim RadiusH As Single
'Dim iCounter As Integer
'Dim R As Single
'Dim Alfa As Single
'
'RadiusW = (cWidth - cLeft) / 2
'RadiusH = (cHeight - cTop) / 2
'
'ReDim POINT(cPtsQty)
'iCounter = 0
'For n = 0 To 360 Step 360 / cPtsQty
'    POINT(iCounter).X = RadiusW + Sin(n * PI / 180) * RadiusW
'    POINT(iCounter).Y = RadiusH + Cos(n * PI / 180) * RadiusH
'    R = Sqr(POINT(iCounter).X ^ 2 + POINT(iCounter).Y ^ 2)
'    Alfa = m2Atn2(POINT(iCounter).Y, POINT(iCounter).X)
'    POINT(iCounter).X = cLeft + R * Cos(Alfa)
'    POINT(iCounter).Y = cTop + R * Sin(Alfa)
'    iCounter = iCounter + 1
'Next
'
'PolygonPoints = POINT
'
'End Function


Public Function PolygonPoints(nPoint As Integer, _
                              cLeft As Single, cTop As Single, cWidth As Single, cHeight As Single, _
                              Optional mAng As Single = 0, Optional mLen As Single) As PointAPI()

Dim POINT() As PointAPI
Dim n As Integer
Dim RadiusW As Single
Dim RadiusH As Single
Dim iCounter As Integer
Dim r As Single
Dim Alfa As Single
       

'    RadiusW = (cWidth - cLeft) / 2
'    RadiusH = (cHeight - cTop) / 2
'    ReDim POINT(nPoint)
'    iCounter = 0
'    For n = 0 To 360 Step 360 / nPoint
'        POINT(iCounter).X = RadiusW + Sin(n * PI / 180) * RadiusW
'        POINT(iCounter).Y = RadiusH + Cos(n * PI / 180) * RadiusH
'        R = Sqr(POINT(iCounter).X ^ 2 + POINT(iCounter).Y ^ 2)
'        Alfa = m2Atn2(POINT(iCounter).Y, POINT(iCounter).X)
'        POINT(iCounter).X = cLeft + R * Cos(Alfa)
'        POINT(iCounter).Y = cTop + R * Sin(Alfa)
'        iCounter = iCounter + 1
'    Next

Dim InRadiusW As Single
Dim InRadiusH As Single

RadiusW = (cWidth - cLeft) / 2
RadiusH = (cHeight - cTop) / 2
'mAng = 0
If mAng = 0 Then mAng = (360 / (nPoint)) / 2
'mLen = 100
InRadiusW = RadiusW * Cos(PI / nPoint) - mLen
InRadiusH = RadiusH * Cos(PI / nPoint) - mLen

ReDim POINT(nPoint * 2)

For n = 0 To 360 Step 360 / nPoint

    POINT(iCounter).X = RadiusW + Sin(n * PI / 180) * RadiusW
    POINT(iCounter).Y = RadiusH + Cos(n * PI / 180) * RadiusH
     r = Sqr(POINT(iCounter).X ^ 2 + POINT(iCounter).Y ^ 2)
    Alfa = m2Atn2(POINT(iCounter).Y, POINT(iCounter).X)
    POINT(iCounter).X = cLeft + r * Cos(Alfa)
    POINT(iCounter).Y = cTop + r * Sin(Alfa)
    iCounter = iCounter + 2
Next

iCounter = 1
For n = 0 To 360 Step (360 / nPoint)
  If iCounter > nPoint * 2 Then Exit For
  POINT(iCounter).X = (RadiusW + Sin((n + mAng) * PI / 180) * (InRadiusW))
  POINT(iCounter).Y = (RadiusH + Cos((n + mAng) * PI / 180) * (InRadiusH))

  r = Sqr(POINT(iCounter).X ^ 2 + POINT(iCounter).Y ^ 2)
  Alfa = m2Atn2(POINT(iCounter).Y, POINT(iCounter).X)
  POINT(iCounter).X = cLeft + r * Cos(Alfa)
  POINT(iCounter).Y = cTop + r * Sin(Alfa)

  iCounter = iCounter + 2

Next

  PolygonPoints = POINT


End Function


'' Find the distance from the point (x0, y0) to the line passing through (x1, y1) and (x2, y2).
'Function LineDistance(ByVal X0 As Single, ByVal Y0 As Single, _
'                      ByVal X1 As Single, ByVal Y1 As Single, _
'                      ByVal X2 As Single, ByVal Y2 As Single) As Single
'
'' Return the distance of a point from a given line
'' x1,y1 First Vertex
'' x2,y2 Second Vertex
'' x0,y0 Point
'
'    Dim Dx As Double, dy As Double
'
'    If (X1 = X2) Then
'        LineDistance = Abs(X1 - X0)
'    ElseIf Y1 = Y2 Then
'        LineDistance = Abs(Y1 - Y0)
'    Else
'        Dx = X2 - X1
'        dy = Y2 - Y1
'        LineDistance = Abs(dy * X0 - Dx * Y0 + X2 * Y1 - X1 * Y2) / Sqr(Dx * Dx + dy * dy)
'    End If
'
'End Function


' Calculate the middle point of the line(X1,Y1)-(X2,Y2).
Public Function MidPoint(ByVal X1 As Single, ByVal Y1 As Single, _
                         ByVal X2 As Single, ByVal Y2 As Single) As PointAPI
        Dim nX As Single, nY As Single
        nX = (X1 + X2) / 2
        nY = (Y1 + Y2) / 2
        MidPoint.X = nX
        MidPoint.Y = nY
End Function

' Find the distance from the point (x1, y1) to the line passing through (x1, y1) and (x2, y2).
Public Function DistPointToLine(ByVal a As Single, ByVal b As Single, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
Dim vx As Single
Dim vy As Single
Dim t As Single
Dim dX As Single
Dim dy As Single
Dim close_x As Single
Dim close_y As Single
On Error GoTo errnum
   
    vx = X2 - X1
    vy = Y2 - Y1

    ' Find the best t value.
    If (vx = 0) And (vy = 0) Then
        ' The points are the same. There is no segment.
        t = 0
    Else
        ' Calculate the minimal value for t.
        If (vx * vx + vy * vy) <> 0 Then
        t = -((X1 - a) * vx + (Y1 - b) * vy) / (vx * vx + vy * vy)
        End If
    End If

    ' Keep the point on the segment.
    If t < 0# Then
        t = 0#
    ElseIf t > 1# Then
        t = 1#
    End If

    ' Set the return values.
    close_x = X1 + t * vx
    close_y = Y1 + t * vy
    dX = a - close_x
    dy = b - close_y
    DistPointToLine = Sqr(dX * dX + dy * dy)
errnum:
   On Error GoTo 0
End Function

' Calculate the distance between the point and the segment.
Public Function DistToSegment(ByVal pX As Single, ByVal pY As Single, _
                              ByVal X1 As Single, ByVal Y1 As Single, _
                              ByVal X2 As Single, ByVal Y2 As Single, _
                              ByRef near_x As Single, ByRef near_y As Single) As Single
Dim dX As Single
Dim dy As Single
Dim t As Single

    dX = X2 - X1
    dy = Y2 - Y1
    If dX = 0 And dy = 0 Then
        ' It's a point not a line segment.
        dX = pX - X1
        dy = pY - Y1
        near_x = X1
        near_y = Y1
        DistToSegment = Sqr(dX * dX + dy * dy)
        Exit Function
    End If

    ' Calculate the t that minimizes the distance.
    t = ((pX - X1) * dX + (pY - Y1) * dy) / (dX * dX + dy * dy)

    ' See if this represents one of the segment's
    ' end points or a point in the middle.
    If t < 0 Then
        dX = pX - X1
        dy = pY - Y1
        near_x = X1
        near_y = Y1
    ElseIf t > 1 Then
        dX = pX - X2
        dy = pY - Y2
        near_x = X2
        near_y = Y2
    Else
        near_x = X1 + t * dX
        near_y = Y1 + t * dy
        dX = pX - near_x
        dy = pY - near_y
    End If
    DistToSegment = Sqr(dX * dX + dy * dy)
End Function

' Return True if the polygon is at this location.
Public Function PolygonIsAt(ByVal is_closed As Boolean, ByVal X As Single, ByVal Y As Single, Points() As PointAPI) As Boolean
Const HIT_DIST = 10
Dim start_i As Integer
Dim i As Integer
Dim num_points As Integer
Dim X1 As Single
Dim Y1 As Single
Dim X2 As Single
Dim Y2 As Single
Dim Dist As Single

    PolygonIsAt = False
     
    num_points = UBound(Points)
    If is_closed Then
        X2 = Points(num_points).X
        Y2 = Points(num_points).Y
        start_i = 1
    Else
        X2 = Points(1).X
        Y2 = Points(1).Y
        start_i = 2
    End If

    ' Check each segment in the Polyline.
    For i = start_i To num_points
        With Points(i)
            X1 = .X
            Y1 = .Y
        End With
        Dist = DistPointToLine(X, Y, X1, Y1, X2, Y2)
        If Dist <= HIT_DIST Then
            PolygonIsAt = True
            Exit For
        End If
        X2 = X1
        Y2 = Y1
    Next i
End Function

' Return True if the point is inside the object.
Public Function PointIsInPolygon(ByVal X As Single, ByVal Y As Single, Points() As PointAPI) As Boolean
Dim polygon_region As Long

    polygon_region = CreatePolygonRgn(Points(1), UBound(Points), fALTERNATE)
    PointIsInPolygon = PtInRegion(polygon_region, X, Y)
    DeleteObject polygon_region
End Function

 'Return a named token from the string txt.
' Tokens have the form TokenName(TokenValue).
Public Sub GetNamedToken(ByRef txt As String, ByRef token_name As String, ByRef token_value As String)
Dim pos1 As Long
Dim pos2 As Long
Dim open_parens As Long
Dim cH As String

    ' Find the "(".
    pos1 = InStr(txt, "(")
    If pos1 = 0 Then
        ' No "(" found. Return the rest as the token name.
        token_name = Trim$(txt)
        token_value = ""
        txt = ""
        Exit Sub
    End If

    ' Find the corresponding ")". Note that parentheses may be nested.
    open_parens = 1
    pos2 = pos1 + 1
    Do While pos2 <= Len(txt)
        cH = mid$(txt, pos2, 1)
        If cH = "(" Then
            open_parens = open_parens + 1
        ElseIf cH = ")" Then
            open_parens = open_parens - 1
            If open_parens = 0 Then
                ' This is the corresponding ")".
                Exit Do
            End If
        End If
        pos2 = pos2 + 1
    Loop

    ' At this point, pos1 points to the ( and pos2 points to the ).
    token_name = Trim$(Left$(txt, pos1 - 1))
    token_value = Trim$(mid$(txt, pos1 + 1, pos2 - pos1 - 1))
    txt = Trim$(mid$(txt, pos2 + 1))
End Sub

' Replace non-printable characters with spaces.
Public Function RemoveNonPrintables(ByVal txt As String) As String
Dim pos As Integer
Dim cH As String
  On Error Resume Next
    
    For pos = 1 To 32
        txt = Replace(txt, Chr(pos), " ")
    Next
    For pos = 126 To 255
        txt = Replace(txt, Chr(pos), " ")
    Next
    RemoveNonPrintables = txt
End Function

Public Function DrawStdPictureRot(ByVal pic As PictureBox, Pts() As PointAPI, ByVal PWidth As Long, ByVal PHeight As Long, _
                                  ByRef inPicture As StdPicture) As Long
  Dim hDC As Long
  Dim hOldBMP As Long
  Dim PlgPts(0 To 4) As PointAPI
  Dim inDC As Long
   Dim PicWidth As Long, PicHeight As Long
   inDC = pic.hDC
  
  ' Validate input picture
  If (inPicture Is Nothing) Then Exit Function
  If (inPicture.Type <> vbPicTypeBitmap) Then Exit Function

  PicWidth = pic.ScaleX(inPicture.Width, vbHiMetric, vbPixels)
  PicHeight = pic.ScaleY(inPicture.Height, vbHiMetric, vbPixels)

  ' Create temporary DC and select input picture into it
  hDC = CreateCompatibleDC(0&)
  hOldBMP = SelectObject(hDC, inPicture.handle)

  If (hOldBMP) Then    ' Get angle vectors for width and height
    PlgPts(2).X = Pts(1).X
    PlgPts(2).Y = Pts(1).Y
    PlgPts(3).X = Pts(2).X
    PlgPts(3).Y = Pts(2).Y
    PlgPts(4).X = Pts(4).X
    PlgPts(4).Y = Pts(4).Y

    ' Draw rotated image
    DrawStdPictureRot = PlgBlt(inDC, PlgPts(2), hDC, 0, 0, PicWidth, PicHeight, 0&, 0, 0)

    ' De-select Bitmap from DC
    Call SelectObject(hDC, hOldBMP)
  End If

  ' Destroy temporary DC
  Call DeleteDC(hDC)

End Function
'
Public Sub FindNodeCurve(ByVal pX1 As Single, ByVal pY1 As Single, ByVal pX2 As Single, ByVal pY2 As Single, _
                          ByRef cX1 As Single, ByRef cY1 As Single, ByRef cX2 As Single, ByRef cY2 As Single)
            Dim tX1 As Single, tY1 As Single
            
            'tX1 = (X1 + X2) / 2
            'tY1 = (Y1 + Y2) / 2
            tX1 = MidPoint(pX1, pY1, pX2, pY2).X ' tX1, tY1
            tY1 = MidPoint(pX1, pY1, pX2, pY2).Y ' tX1, tY1
            cX1 = MidPoint(pX1, pY1, tX1, tY1).X
            cY1 = MidPoint(pX1, pY1, tX1, tY1).Y
            cX2 = MidPoint(tX1, tY1, pX2, pY2).X  ', cX2, cY2
            cY2 = MidPoint(tX1, tY1, pX2, pY2).Y  ', cX2, cY2
           ' MidPoint pX1, pY1, pX2, pY2, tX1, tY1
          '  MidPoint pX1, pY1, tX1, tY1, cX1, cY1
           ' MidPoint tX1, tY1, pX2, pY2, cX2, cY2
End Sub

'Grid Brush
Public Function GetGridBrush(ByVal qCellSize As Integer, qHdc As Long) As Long

Dim brWhite As Long
Dim brGray As Long
Dim brDC As Long
Dim brBmp As Long

Dim qRect As RECT

brWhite = CreateSolidBrush(vbWhite)
brGray = CreateSolidBrush(RGB(200, 200, 200))

brDC = CreateCompatibleDC(qHdc)
brBmp = CreateCompatibleBitmap(qHdc, 2 * qCellSize, 2 * qCellSize)
SelectObject brDC, brBmp

SetRect qRect, 0, 0, qCellSize, qCellSize
FillRect brDC, qRect, brWhite

SetRect qRect, qCellSize, 0, 2 * qCellSize, qCellSize
FillRect brDC, qRect, brGray

SetRect qRect, 0, qCellSize, qCellSize, 2 * qCellSize
FillRect brDC, qRect, brGray

SetRect qRect, qCellSize, qCellSize, 2 * qCellSize, 2 * qCellSize
FillRect brDC, qRect, brWhite

Dim bmpBr As Long
bmpBr = CreatePatternBrush(brBmp)

DeleteDC brDC
DeleteObject brBmp
DeleteObject brWhite
DeleteObject brGray

GetGridBrush = bmpBr

End Function


Public Function CreateShapedRegion2(ByVal hBitmap As Long, _
                                    Optional destinationDC As Long, _
                                    Optional ByVal transColor As Long = -1, _
                                    Optional returnAntiRegion As Boolean) As Long

'*******************************************************
' FUNCTION RETURNS A HANDLE TO A REGION IF SUCCESSFUL.
' If unsuccessful, function retuns zero.
'*******************************************************
' PARAMETERS
'=============
' hBitmap : handle to a bitmap to be used to create the region
' destinationDC : used by GetDibits API. If not supplied then desktop DC used
' transColor : the transparent color
' returnAntiRegion : If False (default) then the region excluding transparent
'       pixels will be returned.  If True, then the region including only
'       transparent pixels will be returned


' test for required variable first
If hBitmap = 0 Then Exit Function

Dim aStart As Long, aStop As Long ' testing purposes
aStart = GetTickCount()            ' testing purposes

' now ensure hBitmap handle passed is a usable bitmap
Dim bmpInfo As BITMAPINFO
If GetObjectAPI(hBitmap, Len(bmpInfo), bmpInfo) = 0 Then Exit Function

' declare bunch of variables...
Dim srcDC As Long   ' DC to use for GetDibits
Dim rgnRects() As RECT ' array of rectangles comprising region
Dim rectCount As Long ' number of rectangles & used to increment above array
Dim rStart As Long ' pixel that begins a new regional rectangle

Dim X As Long, Y As Long ' loop counters
Dim lScanLines As Long ' used to size the DIB bit array

Dim bDib() As Byte  ' the DIB bit array
Dim bBGR(0 To 3) As Byte ' used to copy long to bytes
Dim tgtColor As Long ' a DIB pixel color
Dim rtnRegion As Long ' region handle returned by this function

On Error GoTo CleanUp
  
' use passed DC if supplied, otherwise use desktop DC
If destinationDC = 0 Then
    srcDC = GetDC(0)
Else
    srcDC = destinationDC
End If
    
' Scans must align on dword boundaries:
lScanLines = (bmpInfo.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
ReDim bDib(0 To lScanLines - 1, 0 To bmpInfo.bmiHeader.biHeight - 1)

' build a DIB header
' DIBs are bottom to top, so by using negative Height
' we will load it top to bottom
With bmpInfo.bmiHeader
   .biSize = Len(bmpInfo.bmiHeader)
   .biBitCount = 24
   .biHeight = -.biHeight
   .biPlanes = 1
   .biCompression = 0 'BI_RGB
   .biSizeImage = lScanLines * .biHeight
End With

' get the image into DIB bits,
' note that biHeight above was changed to negative so we reverse it form here on
Call GetDIBits(srcDC, hBitmap, 0, -bmpInfo.bmiHeader.biHeight, bDib(0, 0), bmpInfo, 0)
    
' now get the transparent color if needed
If transColor < 0 Then
    ' when negative value passed, use top left corner pixel color
    CopyMemory bBGR(0), bDib(0, 0), &H3
Else
    ' 24bit DIBs are stored as BGR vs RGB
    ' convert it now for one color vs converting each BGR pixel to RGB
    bBGR(2) = (transColor And &HFF&)
    bBGR(1) = (transColor And &HFF00&) \ &H100&
    bBGR(0) = (transColor And &HFF0000) \ &H10000
End If
' copy bytes to long
CopyMemory transColor, bBGR(0), &H4
    
With bmpInfo.bmiHeader
 
     ' start with an arbritray number of rectangles
    ReDim rgnRects(0 To .biWidth * 3)
    ' reset flag
    rStart = -1
    
    ' begin pixel by pixel comparisons
    For Y = 0 To Abs(.biHeight) - 1
        For X = 0 To .biWidth - 1
            ' my hack continued: we already saved a long as BGR, now
            ' get the current DIB pixel into a long (BGR also) & compare
            CopyMemory tgtColor, bDib(X * 3, Y), &H3
            
            ' test to see if next pixel is a target color
            If transColor = tgtColor Xor returnAntiRegion Then
                
                If rStart > -1 Then ' we're currently tracking a rectangle,
                                    ' so let's close it
                    ' see if array needs to be resized
                   If rectCount + 1 = UBound(rgnRects) Then _
                       ReDim Preserve rgnRects(0 To UBound(rgnRects) + .biWidth * 3)
                    
                    ' add the rectangle to our array
                    SetRect rgnRects(rectCount + 2), rStart, Y, X, Y + 1
                    rStart = -1 ' reset flag
                    rectCount = rectCount + 1     ' keep track of nr in use
                End If
            
            Else
                ' not a target color
                If rStart = -1 Then rStart = X ' set start point
            
            End If
        Next X
        If rStart > -1 Then
            ' got to end of bitmap without hitting another transparent pixel
            ' but we're tracking so we'll close rectangle now
           
                ' see if array needs to be resized
           If rectCount + 1 = UBound(rgnRects) Then _
               ReDim Preserve rgnRects(0 To UBound(rgnRects) + .biWidth * 3)
                ' add the rectangle to our array
            SetRect rgnRects(rectCount + 2), rStart, Y, X, Y + 1
            rStart = -1 ' reset flag
            rectCount = rectCount + 1     ' keep track of nr in use
        End If
    Next Y
End With
Erase bDib
        
On Error Resume Next
' check for failure & engage backup plan if needed
If rectCount Then
    ' there were rectangles identified, try to create the region
    rtnRegion = CreatePartialRegion(rgnRects(), 2, rectCount + 1, 0, bmpInfo.bmiHeader.biWidth)
    
    ' ok, now to test whether or not we are good to go...
    ' if less than 2000 rectangles, API should have worked & if it didn't
    ' it wasn't due O/S restrictions -- failure
    
'    If rtnRegion = 0 And rectCount > 2000 Then
'        rtnRegion = CreateWin98Region(rgnRects, rectCount + 1, 0, bmpInfo.bmiHeader.biWidth)
'    End If

End If

CleanUp:

If destinationDC <> srcDC Then ReleaseDC 0, srcDC
Erase rgnRects()

If Err Then
    If rtnRegion Then DeleteObject rtnRegion
    Err.Clear
    MsgBox "Shaped Region failed. Windows could not create the region."
Else
    CreateShapedRegion2 = rtnRegion
    aStop = GetTickCount()          ' testing purposes
'    MsgBox aStop - aStart & " ms"  ' unRem to show message box
End If


End Function

Private Function CreatePartialRegion(rgnRects() As RECT, lIndex As Long, uIndex As Long, leftOffset As Long, cx As Long) As Long
' Called when large region fails (can be the case with Win98) and also called
' when rotation a region 90 or 270 degrees (see RotateSimpleRegion)

On Error Resume Next
' Note: Ideally contiguous rectangles of equal height & width should be combined
' into one larger rectangle. However, thru trial & error I found that Windows
' does this for us and taking the extra time to do it ourselves
' is to cumbersome & slows down the results.

' the first 32 bytes of a region is the header describing the region.
' Well 32 bytes equates to 2 rectangles (16 bytes each), so I'll
' cheat a little & use rectangles to store the header
With rgnRects(lIndex - 2) ' bytes 0-15
    .Left = 32                      ' length of region header in bytes
    .Top = 1                        ' required cannot be anything else
    .Right = uIndex - lIndex + 1    ' number of rectangles for the region
    .Bottom = .Right * 16&          ' byte size used by the rectangles; can be zero
End With
With rgnRects(lIndex - 1) ' bytes 16-31 bounding rectangle identification
    .Left = leftOffset                  ' left
    .Top = rgnRects(lIndex).Top         ' top
    .Right = leftOffset + cx            ' right
    .Bottom = rgnRects(uIndex).Bottom   ' bottom
End With
' call function to create region from our byte (RECT) array
CreatePartialRegion = ExtCreateRegion(ByVal 0&, (rgnRects(lIndex - 2).Right + 2) * 16, rgnRects(lIndex - 2))
If Err Then Err.Clear
End Function


Public Function StretchRegion(ByVal hSrcRgn As Long, ByVal xScale As Single, ByVal yScale As Single) As Long

' Routine will stretch a region similar to how StretchBlt stretches bitmaps

' hSrcRgn is the region to be stretched
' xScale is percentage of increase or decrease in width
' yScale is percentage of increase or decrease in height
' (i.e., 1.5 for 50% increase and 0.5 for 50% decrease)

' One final note here. I did not take the time to modify my routines
' to handle a failed region in Win98 where the number of rectangles
' may exceed 4,000. If you plan on using this routine, strongly suggest
' modifying the CreateWin98Region & CreatePartialRegion functions to
' accept an XFORM parameter so you can create the region in pieces.
' Examples on how to call the CreateWin98Region are sprinkled throughout
' these routines.

If hSrcRgn = 0 Then Exit Function

Dim xRgn As Long, hBrush As Long, r As RECT

'// ' these are the only UDT members you can set and have function
'   compatible with Win98/Me

    Dim xFrm As XForm
    With xFrm
        .eDx = 0
        .eDy = 0
        .eM11 = xScale
        .eM12 = 0
        .eM21 = 0
        .eM22 = yScale
    End With
    Dim hRgn As Long, dwCount As Long, pRgnData() As Byte
        
    ' get size of region to stretch
    dwCount = GetRegionData(hSrcRgn, 0, ByVal 0&)
    ' create a byte struction to hold that data and get that data
    ReDim pRgnData(0 To dwCount - 1) As Byte
    If dwCount = GetRegionData(hSrcRgn, dwCount, pRgnData(0)) Then
        ' create the stretched region
        hRgn = ExtCreateRegion(xFrm, dwCount, pRgnData(0))
        Erase pRgnData
    End If
StretchRegion = hRgn
End Function

Public Function MaxSelectPointPolygon(pType() As Byte) As Integer
       Dim P As Integer
       For P = 2 To UBound(pType)
           If pType(P) = 2 Then
              MaxSelectPointPolygon = P
              If pType(P - 1) = 4 Then MaxSelectPointPolygon = MaxSelectPointPolygon - 1
              Exit Function
           ElseIf pType(P) = 3 Then
               MaxSelectPointPolygon = 0
                Exit Function
           End If
       Next
End Function

Public Function CreatePolygon(ByVal cx As Single, ByVal cy As Single, _
                              ByVal lX As Single, ByVal lY As Single, _
                              ByVal NumPoints As Integer) As PointAPI()
     Dim phai As Single, theta As Single, hyp As Single, mu As Single
      Dim X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, i As Long
      Dim nPoint() As PointAPI
      ReDim nPoint(1 To NumPoints) As PointAPI
        
        phai = PI / (NumPoints)
        theta = PI / 2 - phai
        hyp = Sqr((lX - cx) * (lX - cx) + (lY - cy) * (lY - cy))
        '//hyp = hypotnuse
        '// erase the previous drawn polygon
        nPoint(1).X = lX
        nPoint(1).Y = lY
        For i = 2 To NumPoints
            X1 = hyp * Cos(theta)
            mu = m2Atn2(lY - cy, lX - cx)
            X2 = X1 * Cos((mu - theta))
            Y2 = X1 * Sin((mu - theta))
            lX = lX - 2 * X2
            lY = lY - 2 * Y2
            nPoint(i).X = lX
            nPoint(i).Y = lY
        Next
      CreatePolygon = nPoint
End Function

's as first point
'p as second point
'cX,cY as center polygon
Public Function CreatePolygon2(ByVal pX1 As Single, ByVal pY1 As Single, _
                               ByVal pX2 As Single, ByVal pY2 As Single, _
                               ByVal NumPoints As Long, _
                               Optional cx As Single, Optional cy As Single) As PointAPI()
                                    
     Dim phai As Single, theta As Single, hyp As Single, mu As Single
     Dim X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, i As Long
     Dim nPoint() As PointAPI
     ReDim nPoint(1 To NumPoints) As PointAPI
     
     Dim Side As Single, LRadius As Single, Inradius As Single, AngOrig As Single
     Dim P1 As PointAPI, P2 As PointAPI, P3 As PointAPI, mX As Single
     Dim mY1 As Single, mX1 As Single
       
     Side = Dist(pX1, pY1, pX2, pY2)
     LRadius = Side / (2 * Sin(PI / NumPoints))
     Inradius = Side / (2 * Tan(PI / NumPoints))
     
     P3 = Circle_intersection(pX1, pY1, LRadius, pX2, pY2, LRadius)
     'Debug.Print "cX:" + Str(P3.X) + ", cY:" + Str(P3.Y)
     nPoint = CreatePolygon(P3.X, P3.Y, pX1, pY1, NumPoints)
     cx = P3.X
     cy = P3.Y
     CreatePolygon2 = nPoint

End Function

Private Function Circle_intersection(x0 As Single, y0 As Single, r0 As Single, _
                       X1 As Single, Y1 As Single, r1 As Single) As PointAPI ', _
                       xi As Single, yi As Single, _
                       xi_prime As Single, yi_prime As Single) As POINTAPIs

                               

  Dim a As Single, dX As Single, dy As Single, d As Single, H As Single, rx As Single, ry As Single
  Dim X2 As Single, Y2 As Single, xi As Single, yi As Single

  '/* dx and dy are the vertical and horizontal distances between
  ' * the circle centers.
  
  dX = X1 - x0
  dy = Y1 - y0

  '/* Determine the straight-line distance between the centers. */
  '//d = sqrt((dy*dy) + (dx*dx));
  d = Sqr((dy * dy) + (dX * dX)) 'hypot(dx, dy) '; // Suggested by Keith Briggs

  '/* Check for solvability. */
  If (d > (r0 + r1)) Or (d < Abs(r0 - r1)) Then
     Exit Function
  End If
  '{
    '/* no solution. circles do not intersect. */
    'return 0
  '}
  
'  If (d < fabs(r0 - r1)) Then
'  '{
'  '  /* no solution. one circle is contained in the other */
'  '  return 0;
'  '}
'
'  /* 'point 2' is the point where the line through the circle
'   * intersection points crosses the line between the circle
'   * centers.
'   */
'
'  /* Determine the distance from point 0 to point 2. */
  a = ((r0 * r0) - (r1 * r1) + (d * d)) / (2# * d)

  '/* Determine the coordinates of point 2. */
  X2 = x0 + (dX * a / d)
  Y2 = y0 + (dy * a / d)

  '/* Determine the distance from point 2 to either of the
  ' * intersection points.
   
  H = Sqr((r0 * r0) - (a * a)) ';

  '/* Now determine the offsets of the intersection points from
  ' * point 2.
  
  rx = -dy * (H / d) ';
  ry = dX * (H / d) ';

  '/* Determine the absolute intersection points. */
  xi = X2 + rx '
  'xi_prime = x2 - rx ';
  yi = Y2 + ry ';
  'yi_prime = y2 - ry ';
  Circle_intersection.X = xi
  Circle_intersection.Y = yi
End Function
  

'Rotating a Point
'Private Function Rotate(P As POINTAPI, Angle As Single) As POINTAPI
Private Function Rotate(ByVal X As Single, ByVal Y As Single, ByVal Angle As Single) As POINTAPIs
' Rotate a single Point using Rad function to converts Degree to Radians
   
    Dim XA As Double, YA As Double
    Dim Seno As Double, Coseno As Double
    Dim P As POINTAPIs
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
Private Function Max(ByVal X1 As Single, ByVal X2 As Single) As Single
     If X1 >= X2 Then
        Max = X1
     Else
        Max = X2
     End If
End Function


Private Function Min(ByVal X1 As Single, ByVal X2 As Single) As Single
     If X1 < X2 Then
        Min = X1
     Else
        Min = X2
     End If
End Function
