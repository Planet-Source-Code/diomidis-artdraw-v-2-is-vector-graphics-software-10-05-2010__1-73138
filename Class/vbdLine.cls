VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' VbDraw Line/Rectangle object.

Implements vbdObject

' Indicates a box rather than a line.
Public IsBox As Boolean
' Constituent vbdPolygon object.
Private m_Polygon As vbdDraw
Private m_Object As vbdObject
Private PointCoords() As PointAPI
Private PointType() As Byte

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
Private m_NumPoints As Integer
Private m_TypePoint() As Byte
Private m_OriginalPoints() As PointAPI
Private m_TransformedPoints() As PointAPI
Private i As Integer
Private IsDrawPoint As Boolean
' Rubberband variables.
Private m_StartX As Single
Private m_StartY As Single
Private m_LastX As Single
Private m_LastY As Single

Private Sub DrawPoint(Optional is_Draw As Boolean = False)
   Dim xmin As Single, ymin As Single, xmax As Single, ymax As Single
   Dim OldScale As typScaleMode, mStep As Single
  
   OldScale = SetWordScale(m_Canvas)
   Polygon m_Canvas.hDC, m_OriginalPoints(1), m_NumPoints
   ResetWordScale m_Canvas, OldScale
    
    mStep = IIf(GAP / gZoomFactor > 0.5, GAP / gZoomFactor, 1)
        m_OriginalPoints(m_NumPoints).X = m_OriginalPoints(1).X
        m_OriginalPoints(m_NumPoints).Y = m_OriginalPoints(1).Y
       For i = 1 To m_NumPoints
           If m_SelectPoint <> i Then
               m_Canvas.Line (m_OriginalPoints(i).X - mStep, m_OriginalPoints(i).Y - mStep)-(m_OriginalPoints(i).X + mStep, m_OriginalPoints(i).Y + mStep), vbGreen, BF
           ElseIf m_SelectPoint = i Then
               m_Canvas.Line (m_OriginalPoints(i).X - mStep, m_OriginalPoints(i).Y - mStep)-(m_OriginalPoints(i).X + mStep, m_OriginalPoints(i).Y + mStep), vbRed, BF
           Else
               m_Canvas.Line (m_OriginalPoints(i).X - mStep, m_OriginalPoints(i).Y - mStep)-(m_OriginalPoints(i).X + mStep, m_OriginalPoints(i).Y + mStep), , BF
           End If
           m_Canvas.Line (m_OriginalPoints(i).X - mStep, m_OriginalPoints(i).Y - mStep)-(m_OriginalPoints(i).X + mStep, m_OriginalPoints(i).Y + mStep), , B
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
           
       End If
End Sub

Private Sub PolyPoints(nPoint As Integer, cx As Single, cy As Single)

   If IsDrawPoint Then
       m_OriginalPoints(nPoint).X = cx '/ gZoomFactor
       m_OriginalPoints(nPoint).Y = cy '/ gZoomFactor
    Else
    If nPoint > 0 Then
        Select Case nPoint
        Case 1
           m_OriginalPoints(1).X = cx '/ gZoomFactor
           m_OriginalPoints(1).Y = cy '/ gZoomFactor
           m_OriginalPoints(2).Y = cy '/ gZoomFactor
           m_OriginalPoints(4).X = cx '/ gZoomFactor
        Case 2
           m_OriginalPoints(2).X = cx '/ gZoomFactor
           m_OriginalPoints(2).Y = cy '/ gZoomFactor
           m_OriginalPoints(1).Y = cy '/ gZoomFactor
           m_OriginalPoints(3).X = cx '/ gZoomFactor
        Case 3
           m_OriginalPoints(3).X = cx '/ gZoomFactor
           m_OriginalPoints(3).Y = cy '/ gZoomFactor
           m_OriginalPoints(2).X = cx '/ gZoomFactor
           m_OriginalPoints(4).Y = cy '/ gZoomFactor
        Case 4
           m_OriginalPoints(4).X = cx '/ gZoomFactor
           m_OriginalPoints(4).Y = cy '/ gZoomFactor
           m_OriginalPoints(1).X = cx '/ gZoomFactor
           m_OriginalPoints(3).Y = cy '/ gZoomFactor
      End Select
    End If
   End If
End Sub

Private Sub m_Canvas_LostFocus()
Dim n As Integer
Dim iCounter As Long
Dim r As Single
Dim Alfa As Single
Dim nfonts As New StdFont
Dim MX As Single, mY As Single
    m_Canvas.DrawMode = vbCopyPen
       
   
    ' Create the vbdPolygon that represents us.
    Set m_Polygon = New vbdDraw
    Set m_Object = m_Polygon
    m_Object.CurrentX = m_StartX
    m_Object.CurrentY = m_StartY
    
    IsDrawPoint = True
    
    With m_Polygon
         .NumPoints = m_NumPoints
          For n = 1 To .NumPoints
               MX = m_OriginalPoints(n).X
               mY = m_OriginalPoints(n).Y
              'PolyPoints n, mX, mY
              .X(n) = MX '/ gZoomFactor
              .Y(n) = mY '/ gZoomFactor
              .TypePoint(n) = 2
          Next
          .TypePoint(1) = 6
          .TypePoint(.NumPoints) = 3
          .IsClosed = True
         
    End With

     If m_TypeDraw = dPicture Then
        Set Me.vbdObject_Picture = LoadPictureGDIPlus(m_Pattern)
     End If
     Me.vbdObject_DrawStyle = m_DrawStyle
     Me.vbdObject_DrawWidth = m_DrawWidth
     Me.vbdObject_FillColor = m_FillColor
     Me.vbdObject_FillColor2 = m_FillColor2
     Me.vbdObject_FillMode = m_FillMode
     Me.vbdObject_FillStyle = m_FillStyle
     Me.vbdObject_ForeColor = m_ForeColor
     Me.vbdObject_TextDraw = m_TextDraw
     Me.vbdObject_TypeDraw = m_TypeDraw
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
     Me.vbdObject_MakeTransformation
     
      ' Stop receiving events from the canvas.
    Set m_Canvas = Nothing
    ' Tell the form to save us.
    AddObject Me
End Sub

' Start drawing a rubberband box.
Private Sub m_Canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If m_DrawingEdit = False Then

        m_DrawingStarted = True

        ' Start using dotted vbInvert mode.
        m_Canvas.DrawMode = vbInvert
        m_Canvas.DrawStyle = vbDot

        ' Start the first rubberband box.
        m_StartX = X '* gZoomFactor
        m_StartY = Y '* gZoomFactor
        m_LastX = X '* gZoomFactor
        m_LastY = Y '* gZoomFactor
        
        If IsBox Then
            m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
        Else
            m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
        End If
    Else
        m_SelectPoint = 0
         For i = 1 To m_NumPoints
            If X >= m_OriginalPoints(i).X - 2 And X <= m_OriginalPoints(i).X + 2 And Y >= m_OriginalPoints(i).Y - 2 And Y <= m_OriginalPoints(i).Y + 2 Then
               m_SelectPoint = i
               Exit For
            End If
        Next
         If m_SelectPoint > 0 Then
           'Start using dotted vbInvert mode.
           m_Canvas.DrawMode = vbInvert
           m_Canvas.DrawStyle = vbDot
           
           DrawPoint
           m_DrawingMove = True
         Else
           'End Draw
           m_Canvas_LostFocus
         End If
    End If
End Sub

' Continue drawing the rubberband box.
Private Sub m_Canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'If Not m_DrawingStarted Then Exit Sub
   If m_DrawingStarted Then
        ' Erase the old box.
        
        If IsBox Then
            m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
        Else
            m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
        End If

        ' Update the point.
        m_LastX = X '* gZoomFactor
        If Shift = 2 Then
            m_LastY = m_StartY + Abs(m_LastX - m_StartX)
        Else
            m_LastY = Y '* gZoomFactor
        End If
    
        ' Draw the new box.
        
        If IsBox Then
            m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
        Else
            m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
        End If
    
    ElseIf m_DrawingEdit Then
        If m_DrawingMove = False Then Exit Sub
        If m_SelectPoint = 0 Then Exit Sub
        'PointCoords = PolyPoints(m_SelectPoint, m_StartX, m_StartY, m_LastX, m_LastY)
        DrawPoint
        ' Update the point.
        m_LastX = X '* gZoomFactor
        If Shift = 2 Then
            m_LastY = m_StartY + Abs(m_LastX - m_StartX)
        Else
            m_LastY = Y '* gZoomFactor
        End If
        
        Call PolyPoints(m_SelectPoint, m_LastX, m_LastY)
        DrawPoint
              
   End If
   
End Sub

' Finish drawing the rubberband box.
Private Sub m_Canvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'If Not m_DrawingStarted Then Exit Sub
  If m_DrawingStarted Then
        ' Erase the old box.
        
        If IsBox Then
            m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
        Else
            m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
        End If
    
        ' Go back to vbCopyPen drawing mode.
          m_Canvas.DrawMode = vbCopyPen
          m_Canvas.DrawStyle = vbSolid
        m_LastX = X '* gZoomFactor
        If Shift = 2 Then
            m_LastY = m_StartY + Abs(m_LastX - m_StartX)
        Else
            m_LastY = Y '* gZoomFactor
        End If
        
        If IsBox Then
            m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
        Else
            m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
        End If
         m_NumPoints = 5
        ReDim m_OriginalPoints(1 To m_NumPoints)
           m_OriginalPoints(1).X = m_StartX
           m_OriginalPoints(1).Y = m_StartY
           m_OriginalPoints(2).X = m_LastX
           m_OriginalPoints(2).Y = m_StartY
           m_OriginalPoints(3).X = m_LastX
           m_OriginalPoints(3).Y = m_LastY
           m_OriginalPoints(4).X = m_StartX
           m_OriginalPoints(4).Y = m_LastY
           m_OriginalPoints(5).X = m_StartX
           m_OriginalPoints(5).Y = m_StartY
        
          ' Go back to vbCopyPen drawing mode.
          'm_Canvas.DrawMode = vbCopyPen
           DrawPoint
           m_DrawingStarted = False
           m_DrawingEdit = True
    
    ElseIf m_DrawingEdit Then
    
        If m_DrawingMove = False Then Exit Sub
       
         DrawPoint
         
         ' Go back to vbCopyPen drawing mode.
         m_Canvas.DrawMode = vbCopyPen
         m_Canvas.DrawStyle = vbSolid
        ' Update the point.
        m_LastX = X '* gZoomFactor
        If Shift = 2 Then
           m_LastY = m_StartY + Abs(m_LastX - m_StartX)
        Else
           m_LastY = Y ''* gZoomFactor
        End If
        m_Canvas_Paint
        Call PolyPoints(m_SelectPoint, m_LastX, m_LastY)
        DrawPoint
        'Polyline m_Canvas.hDC, PointCoords(1), UBound(PointCoords)
        'For I = 1 To 4
        '   m_Canvas.Line (PointCoords(I).X - 2, PointCoords(I).Y - 2)-(PointCoords(I).X + 2, PointCoords(I).Y + 2), , B
        'Next
        ' Go back to vbCopyPen drawing mode.
       ' m_Canvas.DrawMode = vbCopyPen
        m_DrawingMove = False
    End If
    If m_TypeDraw = dPicture And m_NumPoints > 0 Then
       m_Canvas_LostFocus
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
     IsDrawPoint = True
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
            Case "IsBox"
                IsBox = CBool(token_value)
            Case "Polyline", "Scribble", "FreePolygon", "Polygon", "PolyDraw", "Picture", "RectAngle", "Ellipse", "Text", "TextFrame", "TextPath", "Scene"
                m_Object.Serialization = token_value
        End Select
    Loop
End Property

' Return a serialization string for the object.
Public Property Get vbdObject_Serialization() As String
Dim txt As String

    txt = txt & "  IsBox(" & Format$(IsBox) & ") "
    txt = txt & m_Object.Serialization
    Select Case Me.m_TypeDraw
    Case dRectAngle
        vbdObject_Serialization = "RectAngle(" & txt & ")"
    Case dPicture, 0
        vbdObject_Serialization = "Picture(" & txt & ")"
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
