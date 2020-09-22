VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' VbDraw Polygon/Polyline object.

Implements vbdObject

Private Enum vbdDrawErrors
    errInvalidIndex = 381   ' Invalid property array index.
End Enum

' Indicates a closed polygon rather than a polyline.
Public IsClosed As Boolean

' The surface on which the user is clicking
' to define the object. This is set only during
' creation of this object.
'Public WithEvents m_Canvas As PictureBox
Private m_DrawingStarted As Boolean
Private m_EditPoint As Boolean
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
'Public m_TextProperty As New StdFont
Public m_Bold As Boolean
Public m_Charset As Integer
Public m_Italic As Boolean
Public m_Name As String
Public m_Size As Integer
Public m_Strikethrough As Boolean
Public m_Underline As Boolean
Public m_Weight As Integer

Public m_angle As Single
Public m_Selected As Boolean
Public m_hRegion As Long
Public m_Picture As StdPicture

' Data variables.
Private m_NumPoints As Long
Private m_TypePoint() As Byte
Private m_OriginalPoints() As PointAPI
Private m_TransformedPoints() As PointAPI
Private m_M(1 To 3, 1 To 3) As Single

' Rubberband variables.
Private m_StartX As Single
Private m_StartY As Single
Private m_LastX As Single
Private m_LastY As Single

Public Property Let TypePoint(ByVal Index As Integer, ByVal new_value As Integer)
       m_TypePoint(Index) = new_value
End Property

' Return the number of points.
Public Property Get NumPoints() As Integer
    NumPoints = m_NumPoints
End Property

' Set the number of points.
Public Property Let NumPoints(ByVal new_value As Integer)
    m_NumPoints = new_value
    If m_NumPoints < 1 Then
        Erase m_OriginalPoints
        Erase m_TypePoint
    Else
        ReDim Preserve m_OriginalPoints(1 To m_NumPoints)
        ReDim Preserve m_TypePoint(1 To m_NumPoints)
    End If
End Property
' Return an X coordinate.
Property Get X(ByVal Index As Integer) As Single
    If (Index < 1) Or (Index > m_NumPoints) Then
        Err.Raise errInvalidIndex, "Draw.X"
    End If

    X = m_OriginalPoints(Index).X
End Property
' Set an X coordinate.
Property Let X(ByVal Index As Integer, ByVal new_value As Single)
    If (Index < 1) Or (Index > NumPoints) Then
       ' Err.Raise errInvalidIndex, "Draw.X"
    End If

    m_OriginalPoints(Index).X = new_value
End Property

' Return a Y coordinate.
Property Get Y(ByVal Index As Integer) As Single
    If (Index < 1) Or (Index > m_NumPoints) Then
        Err.Raise errInvalidIndex, "Draw.Y"
    End If

    Y = m_OriginalPoints(Index).Y
End Property
' Set a Y coordinate.
Property Let Y(ByVal Index As Integer, ByVal new_value As Single)
    If (Index < 1) Or (Index > NumPoints) Then
       ' Err.Raise errInvalidIndex, "Draw.Y"
    End If
    m_OriginalPoints(Index).Y = new_value
End Property

' Start with an identity transformation.
Private Sub Class_Initialize()
    ' Initialize the drawing parameters.
    InitializeDrawingProperties Me
    m2Identity m_M
End Sub

' The user has selected a point.
Private Sub m_Canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
End Sub

' Continue drawing the rubberband line.
Private Sub m_Canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
End Sub


Private Property Get vbdObject_AlingText() As Integer
        vbdObject_AlingText = m_AlingText
End Property

Private Property Let vbdObject_AlingText(ByVal RHS As Integer)
       m_AlingText = RHS
End Property

Private Property Let vbdObject_Angle(ByVal RHS As Single)
        m_angle = RHS
End Property

Private Property Get vbdObject_Angle() As Single
         vbdObject_Angle = m_angle
End Property

Public Property Let vbdObject_Bold(ByVal RHS As Boolean)
       m_Bold = RHS
End Property

Public Property Get vbdObject_Bold() As Boolean
       vbdObject_Bold = m_Bold
End Property

Private Property Set vbdObject_Canvas(ByVal RHS As PictureBox)
    'Set m_Canvas = RHS
End Property

Private Property Get vbdObject_Canvas() As PictureBox
    'Set vbdObject_Canvas = m_Canvas
End Property

Private Property Let vbdObject_Charset(ByVal RHS As Integer)
         m_Charset = RHS
End Property

Private Property Get vbdObject_Charset() As Integer
          vbdObject_Charset = m_Charset
End Property

' Clear the object's transformation.
Private Sub vbdObject_ClearTransformation()
    m2Identity m_M
    m_TransformedPoints = m_OriginalPoints
End Sub

' Add this transformation to the current one.
Private Sub vbdObject_AddTransformation(m() As Single)
    Dim t(1 To 3, 1 To 3) As Single
    m2MatMultiply t, m_M, m
    m2MatCopy m_M, t
    
End Sub

Private Property Let vbdObject_CurrentX(ByVal RHS As Single)

End Property

Private Property Get vbdObject_CurrentX() As Single

End Property

Private Property Let vbdObject_CurrentY(ByVal RHS As Single)

End Property

Private Property Get vbdObject_CurrentY() As Single

End Property

' Draw the object in a metafile.
Private Sub vbdObject_DrawInMetafile(ByVal mf_dc As Long)
Dim xmin As Single
Dim xmax As Single
Dim ymin As Single
Dim ymax As Single
Dim bR As LogBrush
Dim hBrush As Long, lPen As Long, nL As Long, nT As Long
Dim NRECT As RECT, oldfont As Long, NewFont As Long
Dim OldBrush As Long, hOldPen As Long, mDC As Long, mBitmap As Long
Dim imgData As Long, nW As Long, nH As Long

    On Error GoTo CheckPoints
    If UBound(m_TransformedPoints) = 0 Then
       m_TransformedPoints = m_OriginalPoints
    End If
    
    On Error GoTo 0
    ' Make sure we have at least 2 points.
    If m_NumPoints < 2 Then Exit Sub
    
    TransformPoints
    
    If m_ViewSimple Then
        bR.lbColor = Me.m_FillColor
        bR.lbHatch = -1
        bR.lbStyle = 2
        hBrush = CreateBrushIndirect(bR)
        lPen = PenCreate(Me.m_DrawStyle, 1, Me.m_ForeColor)
        StrokePath mf_dc
    Else
        Select Case Me.m_FillStyle
        Case 0  'Solid
            hBrush = CreateSolidBrush(Me.m_FillColor)
        Case 8  'Bitmap pattern
            imgData = 0 'LoadPictureGDIPlus(App.Path + "\Pattern\" + Me.m_Pattern)
            hBrush = CreatePatternBrush(imgData)
        Case 10  'Bitmap
            vbdObject_Bound xmin, ymin, xmax, ymax
            nW = xmax - xmin
            nH = ymax - ymin
            nL = xmin
            nT = ymin
            mDC = CreateCompatibleDC(0)
            imgData = LoadImage(mDC, Me.m_Pattern, 0, nW, nH, &H10)
            SelectObject mDC, imgData
            If nL < 0 Then nL = 0
            If nT < 0 Then nT = 0
            imgData = BitmapFromDC(mDC, nL, nT, xmax, ymax) ', Me.m_angle, R.Right, R.Bottom)
            
            SelectObject mDC, imgData
            hBrush = CreatePatternBrush(imgData)
           
        Case 9  'Gradient
            vbdObject_Bound xmin, ymin, xmax, ymax
            nW = xmax
            nH = ymax
            'Create a device context, compatible with the screen
            mDC = CreateCompatibleDC(GetDC(0))
            'Create a bitmap, compatible with the screen
            mBitmap = CreateCompatibleBitmap(GetDC(0), nW, nH)
            SelectObject mDC, mBitmap
            GradientFillRectDC mDC, xmin - 5, ymin - 5, xmax, ymax, Me.m_FillColor, Me.m_FillColor2, Me.m_Gradient
            hBrush = CreatePatternBrush(mBitmap)
            
        Case Else   'Other Patterns
            bR.lbColor = Me.m_FillColor
            bR.lbHatch = Me.m_FillStyle - 2
            bR.lbStyle = 2
            hBrush = CreateBrushIndirect(bR)
     End Select
         lPen = PenCreate(Me.m_DrawStyle, Me.m_DrawWidth, Me.m_ForeColor)
    End If
    
    StrokePath mf_dc
    OldBrush = SelectObject(mf_dc, hBrush)
    hOldPen = SelectObject(mf_dc, lPen)
       
    'Make Path for Region
     BeginPath mf_dc
     DrawObjectWMF mf_dc  ', True
     EndPath mf_dc
     
     'Delete  old Region
     DeleteObject Me.m_hRegion
     'New Region
     Me.m_hRegion = PathToRegion(mf_dc)

     'Draw
     If m_TypeDraw = dText Or m_TypeDraw = dPolydraw Or m_TypeDraw = dCurve Or m_TypeDraw = dEllipse Then   'Or m_TypeDraw = dText Then  'Or m_TypeDraw = dTextFrame Then
        DrawObjectWMF mf_dc, True
        DrawObjectWMF mf_dc
     ElseIf m_TypeDraw = dTextFrame Then
        DrawObjectWMF mf_dc, True
    Else
        DrawObjectWMF mf_dc, True
     End If
     
       If m_ViewSimple = False Then
              StrokeAndFillPath mf_dc
       End If
       
    'De-select Brush-pen from DC
    Call SelectObject(mf_dc, OldBrush)
    Call SelectObject(mf_dc, hOldPen)
    
    DeleteObject hBrush
    DeleteObject lPen
    DeleteObject imgData
    DeleteObject mBitmap
    DeleteDC mDC
    Exit Sub
    
CheckPoints:
   If Err = 9 Then
      Resume Next
   End If
End Sub

Private Sub vbdObject_DrawPoint()

End Sub

' Draw box on the object.
Private Sub vbdObject_DrawSelect(ByVal pic As PictureBox)

    Dim xmin As Single
    Dim xmax As Single
    Dim ymin As Single
    Dim ymax As Single
    Dim OldDrawWirdth As Integer
    Dim OldDrawStyle As Integer
    Dim OldFillStyle As Integer
    Dim OldFillColor As Long, hBrush As Long, OldBrush As Long, lPen As Long
    Dim olddrawmode As Integer
    Dim OldFORECOLOR As Long
    Dim r As RECT
    Dim OldMode As Long, OldMapMode As Long
    Dim mStep As Integer
    
    mStep = IIf(GAP / gZoomFactor > 0.5, GAP / gZoomFactor, 1)
    
    ' Make sure we have at least 2 points.
    If m_NumPoints < 2 Then Exit Sub
    
    If ViewingPoints() = False Then Exit Sub
        'Backup parameter
        OldDrawWirdth = pic.DrawWidth
        OldDrawStyle = pic.DrawStyle
        OldFillStyle = pic.FillStyle
        OldFillColor = pic.FillColor
        olddrawmode = pic.DrawMode
        OldFORECOLOR = pic.ForeColor
        
        'New parameter
        pic.DrawWidth = 1
        pic.DrawStyle = vbSolid
        pic.FillStyle = vbFSSolid
        pic.FillColor = vbRed
        pic.DrawMode = vbMaskPenNot
        
    ' Highlight if necessary.
    If m_Selected And m_EditPoint = False Then
        If Me.m_TypeDraw = dTextFrame Then
           vbdObject_Bound xmin, ymin, xmax, ymax
        Else
           GetRgnBox Me.m_hRegion, r
           xmin = r.Left
           ymin = r.Top
           xmax = r.Right
           ymax = r.Bottom
           If (r.Left = 0 And r.Top = 0 And r.Right = 0 And r.Bottom = 0) Or m_TypeDraw = dPolydraw Then
              vbdObject_Bound xmin, ymin, xmax, ymax
           End If
        End If
                 
        If m_ObjLock = False Then
           'Top left
           DrawBoxGap pic, xmin - GAP * 2 / gZoomFactor, ymin - GAP * 2 / gZoomFactor, xmin, ymin
           'Botton left
           DrawBoxGap pic, xmin - GAP * 2 / gZoomFactor, ymax, xmin, ymax + GAP * 2 / gZoomFactor
           'Right top
           DrawBoxGap pic, xmax, ymin - GAP * 2 / gZoomFactor, xmax + GAP * 2 / gZoomFactor, ymin
           'Right botton
           DrawBoxGap pic, xmax, ymax, xmax + GAP * 2 / gZoomFactor, ymax + GAP * 2 / gZoomFactor
        Else
           'Top left
           DrawLock pic, xmin - GAP, ymin - GAP
           'Botton left
           DrawLock pic, xmin - GAP, ymax
           'Right top
           DrawLock pic, xmax, ymin - GAP
           'Right botton
           DrawLock pic, xmax, ymax
        End If
        
        If m_SelectedObjects.Count = 1 Then
          If m_ObjLock = False Then
            'center
            DrawX pic, xmin + (xmax - xmin) / 2, ymin + (ymax - ymin) / 2
            If m_TypeDraw <> dTextFrame Then
            'middle top
            DrawArrowH pic, xmin + (xmax - xmin) / 2, ymin - GAP
            'middle right
            DrawArrowV pic, xmax + GAP, ymin + (ymax - ymin) / 2
            'middle bottom
            DrawArrowH pic, xmin + (xmax - xmin) / 2, ymax + GAP
            'middle right
            DrawArrowV pic, xmin - GAP, ymin + (ymax - ymin) / 2
            End If
            'left for rotate
            DrawCircleGap pic, xmax, ymin + (ymax - ymin) / 2, 20 / gZoomFactor
          End If
        End If
              
    ElseIf m_EditPoint And m_Selected = False Then

       Dim aa As Integer, i As Long, Rad As Integer
       Dim e_NumPoints As Long
       
    If m_TypeDraw = dPolygon Then
       e_NumPoints = 2
    ElseIf m_TypeDraw = dRectAngle Then
       e_NumPoints = 4
    Else
       e_NumPoints = m_NumPoints
    End If
    
      pic.FillColor = vbGreen
      pic.DrawMode = vbCopyPen
'      hBrush = CreateSolidBrush(pic.FillColor)
'      OldBrush = SelectObject(pic.hdc, hBrush)
'      PolyDraw pic.hdc, m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
'      lPen = PenCreate(2, 1, Me.m_ForeColor)
      aa = 0
      ''Debug.Print "DrawSelect"
      If mStep - 1 <= 0 Then Rad = 1 Else Rad = mStep
      For i = 1 To m_NumPoints
          If aa = 3 Then aa = 0
          If m_TypePoint(i) = 4 Then
             aa = aa + 1
             If aa = 1 Then
                'Ellipse pic.hdc, m_TransformedPoints(I).X - mStep - 1, m_TransformedPoints(I).Y - mStep - 1, m_TransformedPoints(I).X + mStep + 1, m_TransformedPoints(I).Y + mStep + 1
                If i <= e_NumPoints Then
                pic.Circle (m_TransformedPoints(i).X, m_TransformedPoints(i).Y), Rad, vbGreen
                 pic.Circle (m_TransformedPoints(i).X, m_TransformedPoints(i).Y), Rad
                End If
             ElseIf aa = 3 Then
                ' Ellipse pic.hdc, m_TransformedPoints(I - 1).X - mStep - 1, m_TransformedPoints(I - 1).Y - mStep - 1, m_TransformedPoints(I - 1).X + mStep + 1, m_TransformedPoints(I - 1).Y + mStep + 1
                If i <= e_NumPoints Then
                pic.Circle (m_TransformedPoints(i - 1).X, m_TransformedPoints(i - 1).Y), Rad, vbGreen
                 pic.Circle (m_TransformedPoints(i - 1).X, m_TransformedPoints(i - 1).Y), Rad
                End If
             ElseIf aa = 2 Then
                ' DrawBoxGap pic, m_TransformedPoints(I + 1).X - mStep - 1, m_TransformedPoints(I + 1).Y - mStep - 1, m_TransformedPoints(I + 1).X + mStep + 1, m_TransformedPoints(I + 1).Y + mStep + 1
                If i <= e_NumPoints Then
                pic.Line (m_TransformedPoints(i + 1).X - mStep, m_TransformedPoints(i + 1).Y - mStep)-(m_TransformedPoints(i + 1).X + mStep, m_TransformedPoints(i + 1).Y + mStep), vbGreen, BF
                pic.Line (m_TransformedPoints(i + 1).X - mStep, m_TransformedPoints(i + 1).Y - mStep)-(m_TransformedPoints(i + 1).X + mStep, m_TransformedPoints(i + 1).Y + mStep), , B
                End If
             End If
          Else
             'DrawBoxGap pic, m_TransformedPoints(I).X - mStep - 1, m_TransformedPoints(I).Y - mStep - 1, m_TransformedPoints(I).X + mStep + 1, m_TransformedPoints(I).Y + mStep + 1
             If i <= e_NumPoints Then
             pic.Line (m_TransformedPoints(i).X - mStep, m_TransformedPoints(i).Y - mStep)-(m_TransformedPoints(i).X + mStep, m_TransformedPoints(i).Y + mStep), vbGreen, BF
             pic.Line (m_TransformedPoints(i).X - mStep, m_TransformedPoints(i).Y - mStep)-(m_TransformedPoints(i).X + mStep, m_TransformedPoints(i).Y + mStep), , B
             End If
          End If
       Next

      aa = 0
       pic.DrawStyle = vbDot
      For i = 1 To m_NumPoints
          If aa = 3 Then aa = 0
          If m_TypePoint(i) = 4 Then
             aa = aa + 1
             If aa = 1 Or aa = 3 Then
               ' MoveToEx pic.hdc, m_TransformedPoints(I - 1).X, m_TransformedPoints(I - 1).Y, ByVal 0
               ' LineTo pic.hdc, m_TransformedPoints(I).X, m_TransformedPoints(I).Y
                'DrawBoxGap pic, m_TransformedPoints(I - 1).X, m_TransformedPoints(I - 1).Y, m_TransformedPoints(I).X, m_TransformedPoints(I).Y
                If i - 1 > 0 Then
                 pic.Line (m_TransformedPoints(i - 1).X, m_TransformedPoints(i - 1).Y)-(m_TransformedPoints(i).X, m_TransformedPoints(i).Y) ', vbRed
                 End If
             End If
          End If
       Next
    End If
    
   
      'Default parameter
        
        pic.DrawWidth = OldDrawWirdth
        pic.DrawStyle = OldDrawStyle
        pic.FillStyle = OldFillStyle
        pic.FillColor = OldFillColor
        pic.DrawMode = olddrawmode
        pic.ForeColor = OldFORECOLOR
End Sub

' Return the object's DrawWidth.
Public Property Get vbdObject_DrawWidth() As Integer
    vbdObject_DrawWidth = m_DrawWidth
End Property

' Set the object's DrawWidth.
Public Property Let vbdObject_DrawWidth(ByVal new_value As Integer)
    m_DrawWidth = new_value
End Property

' Return the object's DrawStyle.
Public Property Get vbdObject_DrawStyle() As Integer
    vbdObject_DrawStyle = m_DrawStyle
End Property
' Set the object's DrawStyle.
Public Property Let vbdObject_DrawStyle(ByVal new_value As Integer)
    m_DrawStyle = new_value
End Property

Private Property Let vbdObject_EditPoint(ByVal RHS As Boolean)
        m_EditPoint = RHS
End Property

Private Property Get vbdObject_EditPoint() As Boolean
         vbdObject_EditPoint = m_EditPoint
End Property

Private Property Let vbdObject_FillColor2(ByVal RHS As stdole.OLE_COLOR)
     m_FillColor2 = RHS
End Property

Private Property Get vbdObject_FillColor2() As stdole.OLE_COLOR
    vbdObject_FillColor2 = m_FillColor2
End Property

Private Property Let vbdObject_FillMode(ByVal RHS As FillMode)
     m_FillMode = RHS
End Property

Private Property Get vbdObject_FillMode() As FillMode
        vbdObject_FillMode = m_FillMode
End Property

' Return the object's ForeColor.
Public Property Get vbdObject_ForeColor() As OLE_COLOR
    vbdObject_ForeColor = m_ForeColor
End Property
' Set the object's ForeColor.
Public Property Let vbdObject_ForeColor(ByVal new_value As OLE_COLOR)
    m_ForeColor = new_value
End Property

' Return the object's FillColor.
Public Property Get vbdObject_FillColor() As OLE_COLOR
    vbdObject_FillColor = m_FillColor
End Property
' Set the object's FillColor.
Public Property Let vbdObject_FillColor(ByVal new_value As OLE_COLOR)
    m_FillColor = new_value
End Property

' Return the object's FillStyle.
Public Property Get vbdObject_FillStyle() As Integer
    vbdObject_FillStyle = m_FillStyle
End Property
' Set the object's FillStyle.
Public Property Let vbdObject_FillStyle(ByVal new_value As Integer)
    m_FillStyle = new_value
End Property

' Return this object's bounds.
Public Sub vbdObject_Bound(ByRef xmin As Single, ByRef ymin As Single, ByRef xmax As Single, ByRef ymax As Single)
Dim i As Integer

    If m_NumPoints < 1 Then
        xmin = 0
        xmax = 0
        ymin = 0
        ymax = 0
    Else
        With m_TransformedPoints(1)
            xmin = .X
            xmax = xmin
            ymin = .Y
            ymax = ymin
        End With

        For i = 2 To m_NumPoints
            With m_TransformedPoints(i)
                If xmin > .X Then xmin = .X
                If xmax < .X Then xmax = .X
                If ymin > .Y Then ymin = .Y
                If ymax < .Y Then ymax = .Y
            End With
        Next i
    End If
End Sub


Public Function CreateBrushFromStdPicture(Image As StdPicture) As Long

  Dim uBI As BITMAP
  Dim uBIH As BITMAPINFO 'BitmapInfoHeader
  Dim aBuff() As Byte '-> Packed DIB

  Dim lhDC As Long
  Dim lhOldBmp As Long

    'If (GetObjectType(Image.Handle) = OBJ_BITMAP) Then

        '-- Get image info
        Call GetObjectAPI(Image.handle, Len(uBI), uBI)

        '-- Prepare DIB header and redim. data array
        With uBIH.bmiHeader
            .biSize = Len(uBIH)
            .biPlanes = 1
            .biBitCount = 24
            .biWidth = uBI.bmWidth
            .biHeight = uBI.bmHeight
            .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
        End With
        ReDim aBuff(1 To Len(uBIH) + uBIH.bmiHeader.biSizeImage)

        '-- Create DIB brush...
        lhDC = CreateCompatibleDC(0)
        If (lhDC <> 0) Then
            lhOldBmp = SelectObject(lhDC, Image.handle)

            '-- Build packed DIB:
            Call CopyMemory(aBuff(1), uBIH, Len(uBIH))
            Call GetDIBits(lhDC, Image.handle, 0, uBI.bmHeight, aBuff(Len(uBIH) + 1), uBIH, 0)  'DIB_RGB_COLORS)

            Call SelectObject(lhDC, lhOldBmp)
            Call DeleteDC(lhDC)

            '-- Create brush from packed DIB
            CreateBrushFromStdPicture = CreateDIBPatternBrushPt(aBuff(1), 0) 'DIB_RGB_COLORS)
        End If
    'End If
End Function

' Draw the object on the canvas.
Public Sub vbdObject_Draw(ByVal pic As PictureBox, Optional SetWsc As Boolean = True)

Dim xmin As Single, X1 As Single
Dim xmax As Single, Y1 As Single
Dim ymin As Single
Dim ymax As Single
Dim hRgn As Long, newRgn As Long
Dim OldMode As Long, OldMapMode As Long
Dim bR As LogBrush, rt As Long
Dim hBrush As Long, lPen As Long, ImageFile As String
Dim OldBrush As Long, hOldPen As Long, mDC As Long, tmDC As Long, mBitmap As Long, mBitmap2 As Long, m_Bitmap As IPictureDisp
Dim imgData As StdPicture, nW As Long, nH As Long, nL As Long, nT As Long, imgData1 As Long
Dim flags As Long, r As RECT
Dim lpRect As RECT
Dim xf As XForm
Dim memDC1 As Long, memDC2 As Long

    On Error GoTo CheckPoints
    If m_NumPoints < 2 Then Exit Sub
    If UBound(m_TransformedPoints) = 0 Then
       m_TransformedPoints = m_OriginalPoints
    End If
    
    On Error GoTo 0
    ' Make sure we have at least 2 points.
    If m_NumPoints < 2 Then Exit Sub
    
    TransformPoints
    
    If ViewingPoints() = False Then Exit Sub
    If m_TypePoint(m_NumPoints) = 3 Then Me.IsClosed = True Else Me.IsClosed = False
    
    If m_ViewSimple Then
        bR.lbColor = Me.m_FillColor
        bR.lbHatch = -1
        bR.lbStyle = 2
        hBrush = CreateBrushIndirect(bR)
        lPen = PenCreate(Me.m_DrawStyle, 1, Me.m_ForeColor)
        StrokePath pic.hDC
    Else
      If m_TypePoint(m_NumPoints) = 3 Then
        Select Case Me.m_FillStyle
        Case 0  'Solid
            hBrush = CreateSolidBrush(Me.m_FillColor)
        Case 8  'Bitmap pattern
            ImageFile = App.Path + "\Pattern\" + Me.m_Pattern
            imgData1 = LoadPatternPic(ImageFile)
            hBrush = CreatePatternBrush(imgData1)
        Case 10  'Bitmap
            vbdObject_Bound xmin, ymin, xmax, ymax
            nW = -1 'xmax - xmin
            nH = -1 'ymax - ymin
          '  nL = xmin
          '  nT = ymin
            mDC = CreateCompatibleDC(0)
            If FileExists(Me.m_Pattern) Then
               Set imgData = LoadPictureGDIPlus(Me.m_Pattern, nW, nH)   'LoadImage(mDC, Me.m_Pattern, 0, nW, nH, &H10)
              '  Set imgData = LoadPictureGDIPlus(Me.m_Pattern, nW * gZoomFactor, nH * gZoomFactor) 'LoadImage(mDC, Me.m_Pattern, 0, nW, nH, &H10)
               If imgData Is Nothing Then Exit Sub
               SelectObject mDC, imgData
               If nL < 0 Then nL = 0
               If nT < 0 Then nT = 0
               imgData1 = BitmapFromDC(mDC, nL, nT, nW, nH)   ', Me.m_angle, R.Right, R.Bottom)
                Call SetBrushOrgEx(mDC, xmin, ymin, ByVal 0&)
               SelectObject mDC, imgData1
              
              '  SelectObject mDC, imgData
               hBrush = CreatePatternBrush(imgData1)
               
               'Stop
              ' Call SetBrushOrgEx(mDC, xmin, ymin, ByVal 0&)
            End If
        Case 9  'Gradient
        
         '  Dim cTile As New cDIBTile
           vbdObject_Bound xmin, ymin, xmax, ymax
            nW = xmax - xmin
            nH = ymax - ymin
            'Create a device context, compatible with the screen
            mDC = CreateCompatibleDC(GetDC(0))
            'Create a bitmap, compatible with the screen
            mBitmap = CreateCompatibleBitmap(GetDC(0), nW, nH)
            SelectObject mDC, mBitmap
            GradientFillRectDC mDC, 0, 0, nW, nH, Me.m_FillColor, Me.m_FillColor2, Me.m_Gradient
           ' Set imgData = mDC
           hBrush = CreatePatternBrush(mBitmap)
           
           
           'SelectObject mDC, mBitmap
           ' Set imgData = CreatePicture(mBitmap)


'            Set m_Bitmap = imgData
'           ' SavePicture imgData, App.Path + "\Pat.bmp"
'            'Set imgData = LoadPictureGDIPlus(App.Path + "\Pat.bmp", nW, nH)
'            cTile.SetPattern m_Bitmap
'           ' SelectObject mDC, mBitmap
'            hBrush = cTile.m_hBrush
            
'             vbdObject_Bound xmin, ymin, xmax, ymax
'            nW = xmax - xmin
'            nH = ymax - ymin
'            'Create a device context, compatible with the screen
'            mDC = CreateCompatibleDC(GetDC(0))
'            'Create a bitmap, compatible with the screen
'            mBitmap = CreateCompatibleBitmap(GetDC(0), nW, nH)
'            SelectObject mDC, mBitmap
'            GradientFillRectDC mDC, xmin, ymin, nW, nH, Me.m_FillColor, Me.m_FillColor2, Me.m_Gradient
'            SelectObject mDC, mBitmap
'             Set imgData = CreatePicture(mBitmap)
'               cTile.SetPattern imgData
'                cTile.SetBrush mDC, xmin, ymin, nW, nH
''            hBrush = CreatePatternBrush(mBitmap)
'             hBrush = cTile.m_hBrush
            ' Set cTile = Nothing
            
            
'          vbdObject_Bound xmin, ymin, xmax, ymax
'            nW = xmax - xmin
'            nH = ymax - ymin
''            'Create a device context, compatible with the screen
''            mDC = CreateCompatibleDC(GetDC(0))
''            'Create a bitmap, compatible with the screen
''            mBitmap = CreateCompatibleBitmap(GetDC(0), nW, nH)
''            SelectObject mDC, mBitmap
''            Call SetBrushOrgEx(mDC, xmin, ymin, ByVal 0&)
''            'GradientFillRectDC mDC, xmin - 5, ymin - 5, xmax, ymax, Me.m_FillColor, Me.m_FillColor2, Me.m_Gradient
''            GradientFillRectDC mDC, 0, 0, nW, nH, Me.m_FillColor, Me.m_FillColor2, Me.m_Gradient
''           SelectObject mDC, mBitmap
''            hBrush = CreatePatternBrush(mBitmap)
''           '  Call SetBrushOrgEx(mDC, xmin, ymin, ByVal 0&)
'
'            'CPaintDC dc(this);
'            'CDC memDC1, memDC2;
'            'memDC1.CreateCompatibleDC(&dc);
'            'memDC2.CreateCompatibleDC(&dc);
'            memDC1 = CreateCompatibleDC(GetDC(0))
'            memDC2 = CreateCompatibleDC(GetDC(0))
'
'            '//You only have to select the source CBitmap in a memory DC A CBitmap bitmap;
'            'bitmap.LoadBitmap(IDB_BITMAP1);
'            mBitmap1 = CreateCompatibleBitmap(GetDC(0), nW, nH)
'            SelectObject memDC1, mBitmap1
'
'            mBitmap2 = CreateCompatibleBitmap(GetDC(0), nW, nH)
'            SelectObject memDC2, mBitmap2
'
'            'memDC1.SelectObject(&bitmap);
'            SelectObject memDC1, mBitmap
'
'            mBitmap2 = CreateCompatibleBitmap(GetDC(0), nW, nH)
'            SelectObject tmDC, mBitmap2
'
'            '//Create your new CBitmap with the new desired size.
'            'CBitmap bitmap2;
'            'bitmap2.CreateCompatibleBitmap(&memDC1,50,50);
'
'            '//Select it into a memory DC B.
'            'memDC2.SelectObject(&bitmap2);
'
'            '//StretchBlt from DC A to DC B.
'            'memDC2.StretchBlt(0,0,50,50,&memDC1,0,0,0,0,SRCCOPY);
'            StretchBlt tmDC, xmin, ymin, xmax, ymax, mDC, 0, 0, 0, 0, vbSrcCopy
'            '//output bitmap to screen
'            'dc.BitBlt(0,0,50,50,&memDC2,0,0,SRCCOPY);
'             SelectObject tmDC, mBitmap
'             hBrush = CreatePatternBrush(mBitmap)
'             DeleteObject mBitmap2
'             DeleteDC mDC
'             DeleteDC tmDC
'
''            'vbdObject_Bound xmin, ymin, xmax, ymax
''            GetRgnBox Me.m_hRegion, R
''            xmin = R.Left
''            xmax = R.Right
''            ymin = R.Top
''            ymax = R.Bottom
''            nW = xmax - xmin
''            nH = ymax - ymin
''            'Create a device context, compatible with the screen
''            mDC = CreateCompatibleDC(GetDC(0))
''            'Create a bitmap, compatible with the screen
''           ' mBitmap = CreateCompatibleBitmap(GetDC(0), nW * gZoomFactor, nH * gZoomFactor)
''            mBitmap = CreateCompatibleBitmap(GetDC(0), nW, nH)
''            SelectObject mDC, mBitmap
''
''            GradientFillRectDC mDC, 0, 0, nW, nH, Me.m_FillColor, Me.m_FillColor2, Me.m_Gradient
''          '   GradientFillRectDC mDC, xmin, ymin, xmax, ymax, Me.m_FillColor, Me.m_FillColor2, Me.m_Gradient
''              Call SetBrushOrgEx(mDC, xmin, ymin, ByVal 0&)
''
''             SelectObject mDC, mBitmap
''
''            Set m_Bitmap = ResizeGDIPlus(mBitmap, vbPicTypeBitmap, (xmax - xmin) * gZoomFactor, (ymax - ymin) * gZoomFactor)
''
''            hBrush = CreateBrushFromStdPicture(m_Bitmap)
'''           ' GradientFillRectDC mDC, xmin, ymin, xmax, ymax, Me.m_FillColor, Me.m_FillColor2, Me.m_Gradient
'''           '  pic.PaintPicture m_Bitmap, xmin, ymin, xmax, ymax
'''             hBrush = CreatePatternBrush(m_Bitmap.Handle)
''             SavePicture m_Bitmap, App.Path + "\Pat.bmp"
''''             hOldBrush = SelectObject(mDC, hBrush)
''''             SelectObject mDC, m_Bitmap
'''             'Call SetBrushOrgEx(mDC, 0, 0, ByVal 0&)
'''             Call SetBrushOrgEx(mDC, 0, 0, ByVal 0&)
'''               SelectObject mDC, mBitmap
'''            hBrush = CreatePatternBrush(mBitmap)
'''            ' Call SetBrushOrgEx(mDC, 10, 10, ByVal 0&)
'''
'''             'SavePicture mBitmap, App.Path + "\Pat.bmp"
'''             'SelectObject mDC, mBitmap
'''
''            hRgn = CreateShapedRegion2(mBitmap, mDC)
''            If hRgn Then
''                'newRgn = StretchRegion(hRgn, 10, 10)
''                'Wxmin, Wymin, Wxmax, Wymax
''                 Call SetBrushOrgEx(pic.hdc, xmin, ymin, ByVal 0&)
''                 FillRgn pic.hdc, hRgn, hBrush
''                ' SetWindowRgn mDC, newRgn, True
''                DeleteObject hRgn
''            End If
''            SelectObject mDC, mBitmap
'''            ''GradientFillRectDC mDC, xmin, ymin, xmax, ymax, Me.m_FillColor, Me.m_FillColor2, Me.m_Gradient
'''           ' hBrush = CreatePatternBrush(mBitmap)
'''         '  SelectObject mDC, mBitmap
'''           'Set the bitmap's backmode to transparent
'''           ' SetBkMode mDC, 1
'''
'''           ' Call DrawPicture(mDC, mBitmap, xmin, ymin, xmax, ymax)
'''           '' BitBlt mDC, 0, 0, nW, nH, pic.hdc, xmin, ymin, vbSrcCopy
'
        Case Else   'Other Patterns
            bR.lbColor = Me.m_FillColor
            bR.lbHatch = Me.m_FillStyle - 2
            bR.lbStyle = 2
            hBrush = CreateBrushIndirect(bR)
       End Select
      End If
     lPen = PenCreate(Me.m_DrawStyle, Me.m_DrawWidth, Me.m_ForeColor)
    End If
            
    OldBrush = SelectObject(pic.hDC, hBrush)
    hOldPen = SelectObject(pic.hDC, lPen)
  
    'Make Path for Region
     BeginPath pic.hDC
     DrawObject pic ', True
     EndPath pic.hDC
        
     'Delete old Region
     DeleteObject Me.m_hRegion
     'Make New Region
     If Me.m_FillMode = 0 Then Me.m_FillMode = fALTERNATE
       If m_NumPoints <= 2 Then
          vbdObject_Bound xmin, ymin, xmax, ymax
          lpRect.Left = xmin
          lpRect.Top = ymin
          lpRect.Right = xmax
          lpRect.Bottom = ymax
          Me.m_hRegion = CreateRectRgnIndirect(lpRect)
       Else
          Me.m_hRegion = CreatePolygonRgn(m_TransformedPoints(1), m_NumPoints, Me.m_FillMode)
       End If

'      If Me.m_FillStyle = 9 Then
'          GetRgnBox Me.m_hRegion, R
'          With xf
'                .eM11 = 1
'                .eM12 = 0
'                .eM21 = 0
'                .eM22 = 1
'                .eDx = R.Left
'                .eDy = R.Top
'            End With
'             rt = SetWorldTransform(pic.hdc, xf)
'
'             OldBrush = SelectObject(pic.hdc, hBrush)
'      End If
   
     'Draw
     If m_TypeDraw = dPolydraw Or m_TypeDraw = dCurve Or m_TypeDraw = dEllipse Then
        DrawObject pic, True
        DrawObject pic
     ElseIf m_TypeDraw = dTextFrame Or m_TypeDraw = dText Then
        DrawObject pic, True      '
        DrawObject pic
    ElseIf m_TypeDraw <> dPicture Then
        DrawObject pic, True
        DrawObject pic
     Else
        
     End If
     
    ' If Me.m_FillStyle <> 9 Then
     If m_ViewSimple = False And m_TypeDraw <> dPicture And Me.IsClosed = True Then
         StrokeAndFillPath pic.hDC
        'StrokePath pic.hdc
    ' End If
     ElseIf m_TypeDraw <> dPicture Then
       ' rt = ModifyWorldTransform(pic.hdc, xf, MWT_IDENTITY)
        FillRgn pic.hDC, Me.m_hRegion, hBrush
     End If
     
    'De-select Brush-pen from DC
    Call SelectObject(pic.hDC, OldBrush)
    Call SelectObject(pic.hDC, hOldPen)
     
    DeleteObject hBrush
    DeleteObject lPen
    If Not imgData Is Nothing Then DeleteObject imgData
    DeleteObject mBitmap
    DeleteDC mDC
    
    If ExportTxt = True Then
       'Stop
       If m_ObjLock = False Then
       AddExportPoint pic, m_TransformedPoints, m_TypePoint
       End If
    End If
    Exit Sub
        
CheckPoints:
   If Err = 9 Then
      Resume Next
   End If
End Sub


'Draw
Private Sub DrawObject(ByVal pic As PictureBox, Optional WithPath As Boolean = False)
   Dim xmin As Single, X1 As Single
   Dim xmax As Single, Y1 As Single
   Dim ymin As Single, i As Long
   Dim ymax As Single, lpXform As XForm, OldXForm As XForm
   Dim mX1 As Long, mX2 As Long, mY1 As Long, mY2 As Long
   Dim BackBuffer As clsBackBuffer, mBuffer As Boolean
   Dim mBlend As Byte, OldMode As Long
   Dim PointCoords() As PointAPI, r As RECT
   Dim PointTypes() As Byte, NumPoints As Long
   
   
    SetPolyFillMode pic.hDC, m_FillMode
       
    vbdObject_Bound xmin, ymin, xmax, ymax
    mX1 = xmin
    mX2 = xmax
    mY1 = ymin
    mY2 = ymax
    
    If WithPath = True And (Me.m_Blend > 0 And Me.m_Blend <= 255) And m_TypeDraw <> dTextFrame Then
        ' Create back-buffer
        Set BackBuffer = New clsBackBuffer
        ' Create a snapshot of the current surface big enough for the drawing
        mBuffer = BackBuffer.CreateSnapshop(pic.hDC, mX1, mY1, (mX2 - mX1), (mY2 - mY1))
    End If
    
    'Text
'    If m_TypeDraw = dText Or m_TypeDraw = dTextFrame Then
'       pic.Font.Bold = m_Bold
'        pic.Font.Charset = m_Charset
'         pic.Font.Italic = m_Italic
'          pic.Font.Name = m_Name
'           pic.Font.Size = m_Size '* gZoomFactor
'            pic.Font.Strikethrough = m_Strikethrough
'             pic.Font.Underline = m_Underline
'             pic.Font.Weight = m_Weight
'
          If m_TypeDraw = dText Then
             vbdObject_Bound xmin, ymin, xmax, ymax
              pic.ForeColor = Me.m_ForeColor
              pic.FillColor = Me.m_FillColor
              If WithPath Then BeginPath pic.hDC
              Call GetWorldTransform(pic.hDC, OldXForm)
               lpXform.eM11 = m_M(1, 1)
               lpXform.eM12 = m_M(1, 2)
               lpXform.eM21 = m_M(2, 1)
               lpXform.eM22 = m_M(2, 2)
               lpXform.eDx = m_M(3, 1)
               lpXform.eDy = m_M(3, 2)
               SetWorldTransform pic.hDC, lpXform
            '   CenterText pic, xmin + ((xmax - xmin) / 2), ymin + ((ymax - ymin) / 2), _
                             m_TextDraw, m_Size, , -m_angle * 10, m_Weight, _
                             m_Italic, m_Underline, m_Strikethrough, m_Charset, , , , , m_Name
                             
                              If WithPath Then BeginPath pic.hDC
             'Draw fill
             PolyDraw pic.hDC, m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
             'PolyDrawVB pic.hdc ', m_TransformedPoints(), m_TypePoint(), m_NumPoints
        If WithPath Then EndPath pic.hDC
        If WithPath Then PolyDraw pic.hDC, m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
                             
              If WithPath Then Call EndPath(pic.hDC)
              Call SetWorldTransform(pic.hDC, OldXForm)
'
'          ElseIf m_TypeDraw = dTextFrame Then
'
'               lpXform.eM11 = m_M(1, 1)
'               lpXform.eM12 = m_M(1, 2)
'               lpXform.eM21 = m_M(2, 1)
'               lpXform.eM22 = m_M(2, 2)
'               lpXform.eDx = m_M(3, 1)
'               lpXform.eDy = m_M(3, 2)
'               m_TransformedPoints = m_OriginalPoints
'               xmin = m_TransformedPoints(1).X
'               ymin = m_TransformedPoints(1).Y
'               xmax = IIf(m_OriginalPoints(3).X > m_TransformedPoints(3).X, m_OriginalPoints(3).X, m_TransformedPoints(3).X)
'               ymax = IIf(m_OriginalPoints(3).Y > m_TransformedPoints(3).Y, m_OriginalPoints(3).Y, m_TransformedPoints(3).Y)
'               DrawFontFrame pic.hDC, m_TextDraw, m_angle, Me.m_ForeColor, m_AlingText, xmin, ymin, xmax, ymax, lpXform
          End If
'
'    Else
      If m_TypeDraw = dPolydraw Or m_TypeDraw = dEllipse Or m_TypeDraw = dCurve Then ' Or m_TypeDraw = dText Then
        If WithPath Then BeginPath pic.hDC
             'Draw fill
             PolyDraw pic.hDC, m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
             'PolyDrawVB pic.hdc ', m_TransformedPoints(), m_TypePoint(), m_NumPoints
        If WithPath Then EndPath pic.hDC
        If WithPath Then PolyDraw pic.hDC, m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
        If WithPath Then FillPath pic.hDC
        
      ElseIf m_TypeDraw = dPicture Then
           Polygon pic.hDC, m_TransformedPoints(1), m_NumPoints
           If m_ViewSimple = False Then
           OldMode = SetStretchBltMode(pic.hDC, &H4)
           DrawStdPictureRot pic, m_TransformedPoints, m_Picture.Width, m_Picture.Height, m_Picture
           SetStretchBltMode pic.hDC, OldMode
           End If
      ElseIf m_TypeDraw <> dText Then
        'Polyline - Polygon
        If WithPath Then BeginPath pic.hDC
             'Draw fill
             PolyDraw pic.hDC, m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
             'PolyDrawVB pic.hdc ', m_TransformedPoints(), m_TypePoint(), m_NumPoints
        If WithPath Then EndPath pic.hDC
        If WithPath Then PolyDraw pic.hDC, m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
'     If WithPath Then FillPath pic.hDC
'         If IsClosed Then
'            'Polygon pic.hdc, m_TransformedPoints(1), m_NumPoints
'            PolyDraw pic.hdc, m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
'         Else
'            'Polyline pic.hdc, m_TransformedPoints(1), m_NumPoints
'             PolyDraw pic.hdc, m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
'         End If
      End If
      
'    End If
    
    If WithPath = True And mBuffer And (Me.m_Blend > 0 And Me.m_Blend < 255) Then 'And m_TypeDraw <> dTextFrame Then
       ' Blend back-buffer to foreground buffer
        mBlend = Me.m_Blend
        Call BackBuffer.Draw(pic.hDC, mX1, mY1, mBlend)
        ''Debug.Print "mBlend:" + Str(mBlend)
       ' Destroy back-buffer
        Set BackBuffer = Nothing
    End If
    
End Sub

Sub DrawObjectWMF(ByVal picwm As Long, Optional WithPath As Boolean = False)
  Dim xmin As Single, X1 As Single
   Dim xmax As Single, Y1 As Single
   Dim ymin As Single, i As Long
   Dim ymax As Single, lpXform As XForm
   Dim mX1 As Long, mX2 As Long, mY1 As Long, mY2 As Long
   Dim BackBuffer As clsBackBuffer, mBuffer As Boolean
   Dim mBlend As Byte
   Dim PointCoords() As PointAPI
   Dim PointTypes() As Byte, NumPoints As Long
   
   
    SetPolyFillMode picwm, m_FillMode
    
    vbdObject_Bound xmin, ymin, xmax, ymax
    mX1 = xmin
    mX2 = xmax
    mY1 = ymin
    mY2 = ymax
    
    If WithPath = True And Me.m_Blend <= 255 And m_TypeDraw <> dTextFrame Then
        ' Create back-buffer
        Set BackBuffer = New clsBackBuffer
        ' Create a snapshot of the current surface big enough for the drawing
        mBuffer = BackBuffer.CreateSnapshop(picwm, mX1, mY1, mX2, mY2)
    End If
    
      If m_TypeDraw = dPolydraw Or m_TypeDraw = dEllipse Or m_TypeDraw = dCurve Then
        If WithPath Then BeginPath picwm
             'Draw fill
             PolyDrawVB picwm ', m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
        If WithPath Then EndPath picwm
        If WithPath Then PolyDrawVB picwm ', m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
        If WithPath Then FillPath picwm
      Else
        'Polyline - Polygon
         If IsClosed Then
            Polygon picwm, m_TransformedPoints(1), m_NumPoints
         Else
            Polyline picwm, m_TransformedPoints(1), m_NumPoints
         End If
      End If
    
    If WithPath = True And mBuffer And Me.m_Blend < 255 And m_TypeDraw <> dTextFrame Then
       ' Blend back-buffer to foreground buffer
        mBlend = Me.m_Blend
        Call BackBuffer.Draw(picwm, mX1, mY1, mBlend)
        ''Debug.Print "mBlend:" + Str(mBlend)
       ' Destroy back-buffer
        Set BackBuffer = Nothing
    End If
    
End Sub
' Apply the transformation matrix.
Public Sub TransformPoints()
    Dim i As Integer
    If m_NumPoints = 0 Then Exit Sub
    ReDim m_TransformedPoints(1 To m_NumPoints)
    For i = 1 To m_NumPoints
        With m_OriginalPoints(i)
            m_TransformedPoints(i).X = (.X * m_M(1, 1) + .Y * m_M(2, 1) + m_M(3, 1))
            m_TransformedPoints(i).Y = (.X * m_M(1, 2) + .Y * m_M(2, 2) + m_M(3, 2))
        End With
    Next i
End Sub

'áëëáãÞ ìå ôï region
Private Function ViewingPoints() As Boolean
    Dim r As RECT
    GetRgnBox Me.m_hRegion, r
         
    'Dim I As Integer
    'For I = 1 To m_NumPoints
       
       If r.Left >= Wxmin And r.Right < Wxmax And _
          r.Top >= Wymin And r.Bottom <= Wymax Then
          ViewingPoints = True
          Exit Function
       End If
       
       If r.Left <= Wxmin And r.Right > Wxmax And _
          r.Top <= Wymin And r.Bottom >= Wymax Then
          ViewingPoints = True
          Exit Function
       End If
       
       If (r.Left >= Wxmin Or r.Left < Wxmax) Or _
          (r.Top >= Wymin And r.Top <= Wymax) Or _
          (r.Right >= Wxmin Or r.Right < Wxmax) Or _
          (r.Bottom >= Wymin And r.Bottom <= Wymax) Then
          ViewingPoints = True
          Exit Function
       End If
   'Next I
End Function

Private Property Let vbdObject_Gradient(ByVal RHS As Integer)
      m_Gradient = RHS
End Property

Private Property Get vbdObject_Gradient() As Integer
         vbdObject_Gradient = m_Gradient
End Property

Private Property Let vbdObject_hRegion(ByVal RHS As Long)
          m_hRegion = RHS
End Property

Private Property Get vbdObject_hRegion() As Long
         vbdObject_hRegion = m_hRegion
End Property

Private Function vbdObject_Info() As String
      vbdObject_Info = "Max Points:" + Str(Me.NumPoints)
End Function

Private Property Get vbdObject_Italic() As Boolean
         vbdObject_Italic = m_Italic
End Property

Private Property Let vbdObject_Italic(ByVal RHS As Boolean)
           m_Italic = RHS
End Property

Public Sub vbdObject_MakeTransformation()
       TransformPoints
End Sub

Private Property Get vbdObject_Name() As String
         vbdObject_Name = m_Name
End Property

Private Property Let vbdObject_Name(ByVal RHS As String)
        m_Name = RHS
End Property

Private Sub vbdObject_NewPoint(NumPoints As Long, X() As Single, Y() As Single, TypePoint() As Byte)
         Dim n As Long
         m_NumPoints = NumPoints
         ReDim m_OriginalPoints(1 To m_NumPoints)
         ReDim m_TypePoint(1 To m_NumPoints)
          For n = 1 To NumPoints
             m_OriginalPoints(n).X = X(n)
             m_OriginalPoints(n).Y = Y(n)
             m_TypePoint(n) = TypePoint(n)
          Next
          Identity
End Sub

Private Sub vbdObject_NewTrPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)

End Sub

Private Property Let vbdObject_ObjLock(ByVal RHS As Boolean)
         m_ObjLock = RHS
End Property

Private Property Get vbdObject_ObjLock() As Boolean
         vbdObject_ObjLock = m_ObjLock
End Property

Private Property Let vbdObject_Blend(ByVal RHS As Integer)
          m_Blend = RHS
End Property

Private Property Get vbdObject_Blend() As Integer
         vbdObject_Blend = m_Blend
End Property

Private Property Let vbdObject_Pattern(ByVal RHS As String)
       m_Pattern = RHS
End Property

Private Property Get vbdObject_Pattern() As String
       vbdObject_Pattern = m_Pattern
End Property

Public Property Get vbdObject_Picture() As stdole.Picture
       Set vbdObject_Picture = m_Picture
End Property

Public Property Set vbdObject_Picture(ByVal RHS As stdole.Picture)
       Set m_Picture = RHS
End Property

Private Sub vbdObject_ReadPoint(ByRef NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
       Dim i As Long
       NumPoints = m_NumPoints
       ReDim X(1 To NumPoints)
       ReDim Y(1 To NumPoints)
       ReDim TPoint(1 To NumPoints)
       TPoint = m_TypePoint
       For i = 1 To m_NumPoints
           X(i) = m_OriginalPoints(i).X
           Y(i) = m_OriginalPoints(i).Y
       Next
End Sub

Private Sub vbdObject_ReadTrPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
       Dim i As Long
       NumPoints = m_NumPoints
       ReDim X(1 To NumPoints)
       ReDim Y(1 To NumPoints)
       ReDim TPoint(1 To NumPoints)
       TPoint = m_TypePoint
       TransformPoints
       For i = 1 To m_NumPoints
           X(i) = m_TransformedPoints(i).X
           Y(i) = m_TransformedPoints(i).Y
       Next
       
End Sub

' Set the object's Selected status.
Private Property Let vbdObject_Selected(ByVal RHS As Boolean)
    m_Selected = RHS
End Property
' Return the object's Selected status.
Private Property Get vbdObject_Selected() As Boolean
    vbdObject_Selected = m_Selected
End Property

' Return True if the object is at this location.
Private Function vbdObject_IsAt(ByVal X As Single, ByVal Y As Single) As Boolean
Dim is_at As Boolean
Dim m_Point() As PointAPI
    
    'm_TransformedPoints = m_OriginalPoints
    On Error GoTo CheckPoints
    If UBound(m_TransformedPoints) = 0 Then m_TransformedPoints = m_OriginalPoints
    
    On Error GoTo vbErr
    m_Point = m_TransformedPoints

    is_at = PolygonIsAt(IsClosed, X, Y, m_Point)
    
    If (Not is_at) And IsClosed And (m_FillStyle <> vbFSTransparent) Then
        is_at = PointIsInPolygon(X, Y, m_Point)
    End If

    vbdObject_IsAt = is_at
    
    On Error GoTo 0
    Exit Function
    
CheckPoints:
    If Err = 9 Then
       'm_TransformedPoints = m_OriginalPoints
       Resume Next
    End If
vbErr:
    vbdObject_IsAt = False
    On Error GoTo 0
End Function

' Initialize the object using a serialization string.
' The serialization does not include the
' ObjectType(...) part.
Private Property Let vbdObject_Serialization(ByVal RHS As String)
Dim token_name As String
Dim token_value As String
Dim next_x As Integer
Dim next_y As Integer
Dim NextPoint As Integer

    InitializeDrawingProperties Me
    m2Identity m_M

    ' Read tokens until there are no more.
    Do While Len(RHS) > 0
        ' Read a token.
        GetNamedToken RHS, token_name, token_value
        Select Case token_name
            Case "IsClosed"
                IsClosed = CBool(token_value)
            Case "NumPoints"
                ' This allocates the m_X and m_Y arrays.
                NumPoints = CLng(token_value)
                next_x = 1
                next_y = 1
                NextPoint = 1
            Case "X"
                X(next_x) = CSng(token_value)
                next_x = next_x + 1
            Case "Y"
                Y(next_y) = CSng(token_value)
                next_y = next_y + 1
            Case "P"
                TypePoint(NextPoint) = CSng(token_value)
                NextPoint = NextPoint + 1
            Case "Transformation"
                SetTransformationSerialization token_value, m_M
            Case "Image"
                 If SaveFileEnc(Environment(s_TEMP) + "\Temp.dat", token_value) Then
                    If FileExists(Environment(s_TEMP) + "\Temp.dat") Then
                       Set Me.m_Picture = LoadPictureGDIPlus(Environment(s_TEMP) + "\Temp.dat")
                       Set Me.vbdObject_Picture = Me.m_Picture
                       Kill Environment(s_TEMP) + "\Temp.dat"
                    End If
                 End If
            Case Else
                ReadDrawingPropertySerialization Me, token_name, token_value
        End Select
    Loop
End Property

' Return a serialization string for the object.
Public Property Get vbdObject_Serialization() As String
Dim txt As String
Dim i As Integer

    txt = DrawingPropertySerialization(Me)
    txt = txt & TransformationSerialization(m_M)
    txt = txt & " IsClosed(" & Format$(IsClosed) & ")"
    txt = txt & " NumPoints(" & Format$(NumPoints) & ")"
    For i = 1 To NumPoints
        With m_OriginalPoints(i)
            txt = txt & vbCrLf & "    X(" & Format$(.X) & ")"
            txt = txt & " Y(" & Format$(.Y) & ")"
            txt = txt & " P(" & Format$(m_TypePoint(i)) & ")"
        End With
    Next i
    
    ''Debug.Print vbdObject_TypeDraw
    Select Case Me.m_TypeDraw
    Case dPolygon
        vbdObject_Serialization = "Polygon(" & txt & ")"
    Case dPolyline
        vbdObject_Serialization = "Polyline(" & txt & ")"
    Case dScribble
        vbdObject_Serialization = "Scribble(" & txt & ")"
    Case dFreePolygon
        vbdObject_Serialization = "FreePolygon(" & txt & ")"
    Case dRectAngle
        vbdObject_Serialization = "RectAngle(" & txt & ")"
    Case dEllipse
        vbdObject_Serialization = "Ellipse(" & txt & ")"
    Case dText
        vbdObject_Serialization = "Text(" & txt & ")"
    Case dTextFrame
        vbdObject_Serialization = "TextFrame(" & txt & ")"
    Case dPolydraw
        vbdObject_Serialization = "PolyDraw(PolyDraw(" & txt & "))"
    Case dCurve
        vbdObject_Serialization = "Curve(" & txt & ")"
    Case dPicture
         If Not Me.m_Picture Is Nothing Then
            SavePicture Me.m_Picture, Environment(s_TEMP) + "\STemp.bmp"
            txt = txt & vbCrLf & " Image (" & vbCrLf & OpenFileEnc(Environment(s_TEMP) + "\STemp.bmp") + " )"
            Kill Environment(s_TEMP) + "\STemp.bmp"
         Else
         txt = txt & vbCrLf & " Image (" & vbCrLf & OpenFileEnc(Me.m_Pattern) + " )"
         End If
        vbdObject_Serialization = "Picture(" & txt & ")"
    End Select
    'Stop
    'vbdObject_Serialization = "Polygon(" & Txt & ")"
   ''Debug.Print vbdObject_Serialization
End Property

Public Property Let vbdObject_Shade(ByVal RHS As Boolean)
       m_Shade = RHS
End Property

Public Property Get vbdObject_Shade() As Boolean
      vbdObject_Shade = m_Shade
End Property

Private Property Let vbdObject_Size(ByVal RHS As Integer)
        m_Size = RHS
End Property

Private Property Get vbdObject_Size() As Integer
          vbdObject_Size = m_Size
End Property

Private Sub vbdObject_StartBound(xmin As Single, ymin As Single)
Dim i As Integer

    If m_NumPoints < 1 Then
        xmin = 0
        ymin = 0
    Else
        With m_OriginalPoints(1)
            xmin = .X
            ymin = .Y
        End With
    End If
End Sub

Private Property Let vbdObject_Strikethrough(ByVal RHS As Boolean)
        m_Strikethrough = RHS
End Property

Private Property Get vbdObject_Strikethrough() As Boolean
         vbdObject_Strikethrough = m_Strikethrough
End Property

Private Property Get vbdObject_TextDraw() As String
        vbdObject_TextDraw = m_TextDraw
End Property

Public Property Let vbdObject_TextDraw(ByVal RHS As String)
     m_TextDraw = RHS
End Property

' Set the object's TypeDraw.
Public Property Let vbdObject_TypeDraw(ByVal RHS As DrawType)
     m_TypeDraw = RHS
End Property

' Return the object's TypeDraw.
Public Property Get vbdObject_TypeDraw() As DrawType
       vbdObject_TypeDraw = m_TypeDraw
End Property

Public Property Let vbdObject_TypeFill(ByVal RHS As DrawTypeFill)
       m_TypeFill = RHS
End Property

Public Property Get vbdObject_TypeFill() As DrawTypeFill
       vbdObject_TypeFill = m_TypeFill
End Property

Private Sub DrawBoxGap(pic As Object, ByVal cX1 As Single, ByVal cY1 As Single, ByVal cX2 As Single, ByVal cY2 As Single)
       RectAngleEx pic.hDC, cX1, cY1, cX2, cY2
End Sub

Private Sub DrawLock(pic As Object, cx As Single, cy As Single)
     Dim N_Point() As PointAPI
     Dim NumPoints(0 To 1) As Long
     Dim cX1 As Long, cY1 As Long, StepGap As Integer
     cX1 = cx
     cY1 = cy
     
     If gZoomFactor > 2 Then
        StepGap = 1
      Else
        StepGap = 2
      End If
      
     ReDim N_Point(1 To 6)
     N_Point(1).X = cX1 + 1 * StepGap
     N_Point(1).Y = cY1 + 3 * StepGap
     N_Point(2).X = cX1 + 1 * StepGap
     N_Point(2).Y = cY1 + 2 * StepGap
     N_Point(3).X = cX1 + 2 * StepGap
     N_Point(3).Y = cY1
     N_Point(4).X = cX1 + 4 * StepGap
     N_Point(4).Y = cY1
     N_Point(5).X = cX1 + 5 * StepGap
     N_Point(5).Y = cY1 + 2 * StepGap
     N_Point(6).X = cX1 + 5 * StepGap
     N_Point(6).Y = cY1 + 3 * StepGap
     Polyline pic.hDC, N_Point(1), 6
     
     ReDim N_Point(1 To 8)
     N_Point(1).X = cX1
     N_Point(1).Y = cY1 + 3 * StepGap
     N_Point(2).X = cX1 + 6 * StepGap
     N_Point(2).Y = cY1 + 3 * StepGap
     N_Point(3).X = cX1 + 6 * StepGap
     N_Point(3).Y = cY1 + 6 * StepGap
     N_Point(4).X = cX1
     N_Point(4).Y = cY1 + 6 * StepGap
     NumPoints(0) = 4
     
     N_Point(5).X = cX1 + 2 * StepGap
     N_Point(6).X = cX1 + 4 * StepGap
     N_Point(7).X = cX1 + 4 * StepGap
     N_Point(8).X = cX1 + 2 * StepGap
     
     N_Point(5).Y = cY1 + 4 * StepGap
     N_Point(6).Y = cY1 + 4 * StepGap
     N_Point(7).Y = cY1 + 5 * StepGap
     N_Point(8).Y = cY1 + 5 * StepGap
     NumPoints(1) = 4
     
     PolyPolygon pic.hDC, N_Point(1), NumPoints(0), 2
 
End Sub

Private Sub DrawArrowV(pic As Object, ByVal cx As Long, ByVal cy As Long)
     Dim N_Point(0 To 4) As PointAPI
     Dim cX1 As Long, cY1 As Long, StepGap As Integer
     cX1 = cx
     cY1 = cy
     
     If GAP / gZoomFactor > 1 Then
        StepGap = GAP / gZoomFactor
     Else
        StepGap = 1
     End If
     
     MoveToEx pic.hDC, cX1, cY1 - StepGap, ByVal 0&
     LineTo pic.hDC, cX1, cY1 + StepGap
     
     N_Point(0).Y = cY1 - StepGap * 2
     N_Point(0).X = cX1
     N_Point(1).Y = cY1 - StepGap
     N_Point(1).X = cX1 - StepGap
     N_Point(2).Y = cY1 - StepGap
     N_Point(2).X = cX1 + StepGap
     N_Point(3).Y = cY1 - StepGap * 2
     N_Point(3).X = cX1
     Polygon pic.hDC, N_Point(0), 4

     N_Point(0).Y = cY1 + StepGap * 2
     N_Point(0).X = cX1
     N_Point(1).Y = cY1 + StepGap
     N_Point(1).X = cX1 - StepGap
     N_Point(2).Y = cY1 + StepGap
     N_Point(2).X = cX1 + StepGap
     N_Point(3).Y = cY1 + StepGap * 2
     N_Point(3).X = cX1
     Polygon pic.hDC, N_Point(0), 4

End Sub

Private Sub DrawArrowH(pic As Object, ByVal cx As Long, ByVal cy As Long)
     Dim N_Point(1 To 3) As PointAPI
     Dim cX1 As Long, cY1 As Long, StepGap As Integer
     
     If GAP / gZoomFactor > 1 Then
        StepGap = GAP / gZoomFactor
     Else
        StepGap = 1
     End If
           
     cX1 = cx
     cY1 = cy
     
     MoveToEx pic.hDC, cX1 - StepGap, cY1, ByVal 0&
     LineTo pic.hDC, cX1 + StepGap, cY1
     
     N_Point(1).X = cX1 - StepGap * 2
     N_Point(1).Y = cY1
     N_Point(2).X = cX1 - StepGap
     N_Point(2).Y = cY1 - StepGap
     N_Point(3).X = cX1 - StepGap
     N_Point(3).Y = cY1 + StepGap
     Polygon pic.hDC, N_Point(1), 3

     N_Point(1).X = cX1 + StepGap * 2
     N_Point(1).Y = cY1
     N_Point(2).X = cX1 + StepGap
     N_Point(2).Y = cY1 - StepGap
     N_Point(3).X = cX1 + StepGap
     N_Point(3).Y = cY1 + StepGap
     Polygon pic.hDC, N_Point(1), 3
     
End Sub

Private Sub DrawX(pic As Object, cx As Single, cy As Single)
      Dim OldWidth As Integer, StepGap As Integer

      OldWidth = pic.DrawWidth
     
      If gZoomFactor <= 1 Then
         pic.DrawWidth = (2 / gZoomFactor) * 2
      ElseIf gZoomFactor < 2 Then
         pic.DrawWidth = 4
      ElseIf gZoomFactor > 2 And gZoomFactor <= 4 Then
         pic.DrawWidth = 2
      Else
         pic.DrawWidth = 1
      End If
      If GAP / gZoomFactor > 1 Then
         StepGap = GAP / gZoomFactor
      Else
         StepGap = 1
      End If
      
      MoveToEx pic.hDC, cx - StepGap, cy - StepGap, ByVal 0&
      LineTo pic.hDC, cx + StepGap, cy + StepGap

      MoveToEx pic.hDC, cx - StepGap, cy + StepGap, ByVal 0&
      LineTo pic.hDC, cx + StepGap, cy - StepGap
     
      pic.DrawWidth = OldWidth
End Sub

Private Sub DrawCenter(pic As Object, cx As Single, cy As Single)
      Dim OldWidth As Integer
      OldWidth = pic.DrawWidth
      pic.PSet (cx, cy)
      pic.DrawWidth = 1
      pic.FillColor = vbRed
      Ellipse pic.hDC, cx - 6, cy - 6, cx + 6, cy + 6
      pic.DrawWidth = OldWidth
End Sub

Private Sub DrawCircleGap(pic As Object, cx As Single, cy As Single, Optional cLenght As Integer = 20)
     Dim Ls As Single, StepGap As Integer
     
     If GAP / gZoomFactor > 2 Then
        StepGap = GAP / gZoomFactor
     Else
        StepGap = 2
     End If
     
     If cLenght < 20 Then cLenght = 20
     pic.FillColor = vbRed
     
     MoveToEx pic.hDC, cx + GAP, cy, ByVal 0&
     LineTo pic.hDC, cx + cLenght - StepGap, cy

     pic.FillColor = vbRed
     Ellipse pic.hDC, cx + cLenght - StepGap, cy - StepGap, cx + cLenght + StepGap, cy + StepGap

End Sub

Private Property Let vbdObject_Underline(ByVal RHS As Boolean)
     m_Underline = RHS
End Property

Private Property Get vbdObject_Underline() As Boolean
          vbdObject_Underline = m_Underline
End Property

Private Property Let vbdObject_Weight(ByVal RHS As Integer)
        m_Weight = RHS
End Property

Private Property Get vbdObject_Weight() As Integer
        vbdObject_Weight = m_Weight
End Property

Public Sub Identity()
       m2Identity m_M
End Sub

 
Private Sub PolyDrawVB(ByVal wmhdc As Long)
                       
    Dim LoopPts As Long
    Dim BezIdx As Long
    Dim lpPt() As PointAPI, lpbTypes() As Byte, cCount As Long
    lpPt = m_TransformedPoints '(1)
    lpbTypes = m_TypePoint '(1)
    cCount = m_NumPoints
    Const PT_CLOSEFIGURE As Long = &H1
    Const PT_LINETO As Long = &H2
    Const PT_BEZIERTO As Long = &H4
    Const PT_MOVETO As Long = &H6
    
    For LoopPts = 1 To cCount '- 1
        ' Clear bezier point index for non-bezier point
        If ((lpbTypes(LoopPts) And PT_BEZIERTO) = 0) Then BezIdx = 0

        Select Case lpbTypes(LoopPts) And Not PT_CLOSEFIGURE
            Case PT_LINETO    ' Straight line segment
                Call LineTo(wmhdc, lpPt(LoopPts).X, lpPt(LoopPts).Y)
        Case PT_BEZIERTO    ' Curve segment
                '//
                Select Case BezIdx
                Case 0, 1   ' Bezier control handles
                   '  Call DrawBox(lpPt(LoopPts), 2, vbBlue)
                   ' Me.Caption = LoopPts + 1 & "/" & cCount & " Bezier control handles :" & BezIdx
                Case 2    ' Bezier end point
                   '//Connecting lines betweenn (start to 1st control pt) and (2nd control pt to end point)
                   ' Move to first point where we will start curve
                   Call MoveToEx(wmhdc, lpPt(LoopPts - 3).X, lpPt(LoopPts - 3).Y, ByVal 0&)
                   Call PolyBezierTo(wmhdc, lpPt(LoopPts - 2), 3)
                End Select
                   BezIdx = (BezIdx + 1) Mod 3 '//Reset counter after 3 Bezier points
            Case PT_MOVETO    ' Move current drawing point
                 Call MoveToEx(wmhdc, lpPt(LoopPts).X, lpPt(LoopPts).Y, ByVal 0&)
        End Select

        If (lpbTypes(LoopPts) And PT_CLOSEFIGURE) Then
            Call CloseFigure(wmhdc)
        End If
    Next LoopPts
    
End Sub



