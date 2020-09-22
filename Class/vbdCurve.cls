VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdCurve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Private m_OriginalPoints() As POINTAPI
Private m_TransformedPoints() As POINTAPI
Private I As Integer

' Rubberband variables.
Private m_StartX As Single
Private m_StartY As Single
Private m_LastX As Single
Private m_LastY As Single

Private Sub DrawPoint(Optional is_Draw As Boolean = False)
     Dim oldmode As Integer
     
       PolyDraw m_Canvas.hDC, m_OriginalPoints(1), m_TypePoint(1), m_NumPoints
       oldmode = m_Canvas.DrawStyle
       m_Canvas.DrawStyle = vbDot
       For I = 1 To m_NumPoints Step 2
          If I = 1 And m_NumPoints > 1 Then
             m_Canvas.Line (m_OriginalPoints(I).X, m_OriginalPoints(I).Y)-(m_OriginalPoints(I + 1).X, m_OriginalPoints(I + 1).Y), vbBlue
          ElseIf I + 1 <= m_NumPoints Then
             m_Canvas.Line (m_OriginalPoints(I - 1).X, m_OriginalPoints(I - 1).Y)-(m_OriginalPoints(I).X, m_OriginalPoints(I).Y), vbBlue
             m_Canvas.Line (m_OriginalPoints(I).X, m_OriginalPoints(I).Y)-(m_OriginalPoints(I + 1).X, m_OriginalPoints(I + 1).Y), vbBlue
          End If
       Next
       If m_NumPoints > 1 Then
          m_Canvas.Line (m_OriginalPoints(m_NumPoints - 1).X, m_OriginalPoints(m_NumPoints - 1).Y)-(m_OriginalPoints(m_NumPoints).X, m_OriginalPoints(m_NumPoints).Y), vbBlue
       End If
       m_Canvas.DrawStyle = oldmode
       For I = 1 To m_NumPoints
           If m_SelectPoint <> I Then
               m_Canvas.Line (m_OriginalPoints(I).X - 2, m_OriginalPoints(I).Y - 2)-(m_OriginalPoints(I).X + 2, m_OriginalPoints(I).Y + 2), vbGreen, BF
           ElseIf m_SelectPoint = I Then
               m_Canvas.Line (m_OriginalPoints(I).X - 2, m_OriginalPoints(I).Y - 2)-(m_OriginalPoints(I).X + 2, m_OriginalPoints(I).Y + 2), vbRed, BF
           Else
               m_Canvas.Line (m_OriginalPoints(I).X - 2, m_OriginalPoints(I).Y - 2)-(m_OriginalPoints(I).X + 2, m_OriginalPoints(I).Y + 2), , BF
           End If
           'If m_TypePoint(I) <> 3 Then
           m_Canvas.Line (m_OriginalPoints(I).X - 2, m_OriginalPoints(I).Y - 2)-(m_OriginalPoints(I).X + 2, m_OriginalPoints(I).Y + 2), , B
           'End If
       Next
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
           ' Me.m_ShadeX=Me.vbdObject_ShadeX
           ' Me.m_ShadeY=Me.vbdObject_ShadeY
           ' Me.m_ShadeFillStyle=Me.vbdObject_ShadeFillStyle
           ' Me.m_ShadeGradient=Me.vbdObject_ShadeGradient
           ' Me.m_ShadePattern=Me.vbdObject_ShadePattern
           ' Me.m_Gradient  = Me.vbdObject_Gradient
       End If
End Sub

Private Sub PolyPoints(nPoint As Integer, CX As Single, CY As Single)

    If nPoint > 0 Then
        m_OriginalPoints(nPoint).X = CX '/ gZoomFactor
        m_OriginalPoints(nPoint).Y = CY '/ gZoomFactor
    End If
    'Debug.Print nPoint, m_OriginalPoints(nPoint).X, m_OriginalPoints(nPoint).Y
End Sub

Private Sub m_Canvas_KeyUp(KeyCode As Integer, Shift As Integer)
     Dim Points() As POINTAPI, aa As Long
       If KeyCode = vbKeyDelete Then
          If m_NumPoints > 3 Then
          If m_SelectPoint > 1 And m_SelectPoint <= m_NumPoints Then
            ' If m_TypePoint(m_SelectPoint) <> 3 And m_TypePoint(m_SelectPoint) <> 6 Then
             m_NumPoints = m_NumPoints - 1
             ReDim Points(1 To m_NumPoints)
             aa = 0
             For I = 1 To m_NumPoints + 1
                If m_SelectPoint <> I Then
                  aa = aa + 1
                  Points(aa).X = m_OriginalPoints(I).X
                  Points(aa).Y = m_OriginalPoints(I).Y
                End If
             Next
             m_OriginalPoints = Points
            ' End If
          End If
          End If
       End If
       DrawPoint
End Sub

Private Sub m_Canvas_LostFocus()
  Dim mX As Single, mY As Single
      
       Set m_Polygon = New vbdDraw
       Set m_Object = m_Polygon
       m_DrawingEdit = False
        m_Canvas.DrawMode = vbCopyPen

        ' Stop receiving events from the canvas.
        Set m_Canvas = Nothing

        ' If we have at least 2 points, tell the form to save us.
        If m_NumPoints >= 2 Then
            m_TypePoint(1) = 6
            If IsClosed = True Then m_TypePoint(m_NumPoints) = 3
            
            With m_Polygon
                .NumPoints = m_NumPoints
                For I = 1 To m_NumPoints
                    ' mX = m_OriginalPoints(I).X
                    ' mY = m_OriginalPoints(I).Y
                  ' PolyPoints I, mX, mY
                   .X(I) = m_OriginalPoints(I).X / gZoomFactor
                   .Y(I) = m_OriginalPoints(I).Y / gZoomFactor
                   .TypePoint(I) = m_TypePoint(I)
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
    ' If we are drawing a rubberband line, erase it.
        
    'If Button <> 1 Then Exit Sub
    ' If this is the first point, start using dotted vbInvert mode.
    If m_NumPoints = 0 Then
        m_Canvas.DrawMode = vbInvert
        m_Canvas.DrawStyle = vbDot
        m_DrawingStarted = True
    End If
  
    If m_DrawingEdit = False Then
        If m_NumPoints > 0 Then m_Canvas.Line (X - 2, Y - 2)-(X + 2, Y + 2), vbGreen, B '(m_StartX, m_StartY)-(m_LastX, m_LastY)
       ' See if this is the left or right button.
       If Button And vbLeftButton Then
            ' It's the left button.
            ' Add this point to the polygon.

            ' Add the new point.
            m_NumPoints = m_NumPoints + 1
            ReDim Preserve m_OriginalPoints(1 To m_NumPoints)
            ReDim Preserve m_TypePoint(1 To m_NumPoints)
            
            m_OriginalPoints(m_NumPoints).X = X
            m_OriginalPoints(m_NumPoints).Y = Y
            If m_NumPoints = 1 Then
               m_TypePoint(1) = 6
            Else
               m_TypePoint(m_NumPoints) = 4
            End If
            ' Draw the line permanently.
            m_Canvas.DrawMode = vbCopyPen
            m_Canvas.DrawStyle = vbSolid
            'm_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
            m_Canvas.Line (X - 2, Y - 2)-(X + 2, Y + 2), vbGreen, B
            m_Canvas.DrawMode = vbInvert
            m_Canvas.DrawStyle = vbDot

            ' Start the next rubberband line.
            m_StartX = X '* gZoomFactor
            m_StartY = Y '* gZoomFactor
            m_LastX = X '* gZoomFactor
            m_LastY = Y '* gZoomFactor
            'm_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
            m_Canvas.Line (X - 2, Y - 2)-(X + 2, Y + 2), vbGreen, B
            DrawPoint
       Else
           m_Canvas.DrawMode = vbCopyPen
           m_Canvas.DrawStyle = vbSolid
           
           DrawPoint
           
           m_DrawingStarted = False
           m_DrawingEdit = True
       End If
       
    Else
    
        m_SelectPoint = 0
        For I = 1 To m_NumPoints
            If X >= m_OriginalPoints(I).X - 2 And X <= m_OriginalPoints(I).X + 2 And Y >= m_OriginalPoints(I).Y - 2 And Y <= m_OriginalPoints(I).Y + 2 Then
               m_SelectPoint = I
               Exit For
            End If
        Next
        
        If m_SelectPoint > 0 Then
           m_Canvas.DrawMode = vbInvert
           m_Canvas.DrawStyle = vbDot

           m_LastX = X '/ gZoomFactor
           m_LastY = Y '/ gZoomFactor
           
           DrawPoint
           
           m_DrawingMove = True
        Else
           'End Draw
           m_Canvas_LostFocus
        End If
    End If
End Sub

' Continue drawing the rubberband line.
Private Sub m_Canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 1 Then Exit Sub
    'If Not m_DrawingStarted Then Exit Sub
    If m_DrawingStarted Then

        ' Start using dotted vbInvert mode.
        m_Canvas.DrawMode = vbInvert
        m_Canvas.DrawStyle = vbDot

        ' Erase the old line.
        m_Canvas.Line (X - 2, Y - 2)-(X + 2, Y + 2), vbGreen, B
        'm_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)

        ' Update the point.
        m_LastX = X '* gZoomFactor
        m_LastY = Y '* gZoomFactor

        ' Draw the new line.
        'm_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
        m_Canvas.Line (X - 2, Y - 2)-(X + 2, Y + 2), vbGreen, B
        
    ElseIf m_DrawingEdit Then
        If m_DrawingMove = False Then Exit Sub
                
        ' Start using dotted vbInvert mode.
      '  m_Canvas.DrawMode = vbInvert
        
        'Call PolyPoints(m_SelectPoint, m_LastX, m_LastY)
        DrawPoint
        ' Update the point.
        m_LastX = X '/ gZoomFactor
        m_LastY = Y '/ gZoomFactor

        Call PolyPoints(m_SelectPoint, m_LastX, m_LastY)
        DrawPoint
              
   End If
   
End Sub


Private Sub m_Canvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button <> 1 Then Exit Sub
     If m_DrawingEdit Then
     
        If m_DrawingMove = False Then Exit Sub
        'Call PolyPoints(m_SelectPoint, m_LastX, m_LastY)
        DrawPoint
        ' Go back to vbCopyPen drawing mode.
         m_Canvas.DrawMode = vbCopyPen
         m_Canvas.DrawStyle = vbSolid
        
        m_Canvas_Paint
        
        'Call PolyPoints(m_SelectPoint, X / gZoomFactor, Y / gZoomFactor)
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
Private Sub vbdObject_AddTransformation(M() As Single)
    m_Object.AddTransformation M
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
    m_Canvas_Paint
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
Public Sub vbdObject_Draw(ByVal pic As PictureBox)
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
      Dim n As Long
      m_Polygon.NumPoints = NumPoints
              
      For n = 1 To NumPoints
          m_Polygon.X(n) = X(n) + m_Object.CurrentX
          m_Polygon.Y(n) = Y(n) + m_Object.CurrentY
          m_Polygon.TypePoint(n) = TypePoint(n)
      Next
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
            Case "Polyline", "Scribble", "FreePolygon", "Polygon", "PolyDraw", "Picture", "RectAngle", "Ellipse", "Text", "TextFrame", "TextPath", "Scene", "Curve"
                m_Object.Serialization = token_value
        End Select
    Loop
End Property

' Return a serialization string for the object.
Public Property Get vbdObject_Serialization() As String
Dim txt As String

    txt = txt & " IsClosed(" & Format$(IsClosed) & ")"
    txt = txt & m_Object.Serialization
    vbdObject_Serialization = "Curve(" & txt & ")"
    

End Property

Public Property Let vbdObject_Shade(ByVal RHS As Boolean)
       m_Object.Shade = RHS
End Property

Public Property Get vbdObject_Shade() As Boolean
      vbdObject_Shade = m_Object.Shade
End Property

Public Property Let vbdObject_Size(ByVal RHS As Integer)
        m_Object.size = RHS
End Property

Public Property Get vbdObject_Size() As Integer
          vbdObject_Size = m_Object.size
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

