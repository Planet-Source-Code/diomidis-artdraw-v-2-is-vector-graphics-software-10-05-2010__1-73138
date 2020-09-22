VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdEllipse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' VbDraw Ellipse object.

Implements vbdObject

' Indicates a box rather than a line.
Public IsBox As Boolean

' The surface on which the user is clicking
' to define the object. This is set only during
' creation of this object.
Public WithEvents m_Canvas As PictureBox
Attribute m_Canvas.VB_VarHelpID = -1
Private m_DrawingStarted As Boolean
Public m_DrawingEdit As Boolean
Private m_DrawingMove As Boolean
Private m_SelectPoint As Integer

' Constituent vbdPolygon object.
Private m_Polygon As vbdDraw
Private m_Object As vbdObject

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
' Rubberband variables.
Private m_StartX As Single
Private m_StartY As Single
Private m_LastX As Single
Private m_LastY As Single
Public m_MovePoints As Boolean
Public m_MaxPoints As Integer

'' Data variables.
Private m_NumPoints As Integer
Private m_TypePoint() As Byte
Private m_OriginalPoints() As PointAPI
Private m_TransformedPoints() As PointAPI
Private i As Integer


Private Sub Class_Initialize()
     m_MaxPoints = 10
End Sub

' Start drawing a rubberband box.
Private Sub m_Canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_DrawingStarted = True

    ' Start using dotted vbInvert mode.
    m_Canvas.DrawMode = vbInvert
    m_Canvas.DrawStyle = vbDot
    'm_Canvas.AutoRedraw = False
    ' Start the first rubberband box.
    m_StartX = X
    m_StartY = Y
    m_LastX = X
    m_LastY = Y
    m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
    If (m_LastX <> m_StartX) Then
       m_Canvas.Circle ((m_StartX + m_LastX) / 2, (m_StartY + m_LastY) / 2), varIIf(Abs(m_LastX - m_StartX) > Abs(m_LastY - m_StartY), Abs(m_LastX - m_StartX) / 2, Abs(m_LastY - m_StartY) / 2), , , , Abs((m_LastY - m_StartY) / (m_LastX - m_StartX))
    End If
    
End Sub

' Continue drawing the rubberband box.
Private Sub m_Canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
Dim cx As Single
Dim cy As Single
Dim rx As Single
Dim ry As Single
Dim ptx() As Single
Dim pty() As Single
Dim i As Integer
Dim dtheta As Single

    If Not m_DrawingStarted Then Exit Sub
     
    ' Erase the old.
     m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
     If (m_LastX <> m_StartX) Then
       m_Canvas.Circle ((m_StartX + m_LastX) / 2, (m_StartY + m_LastY) / 2), varIIf(Abs(m_LastX - m_StartX) > Abs(m_LastY - m_StartY), Abs(m_LastX - m_StartX) / 2, Abs(m_LastY - m_StartY) / 2), , , , Abs((m_LastY - m_StartY) / (m_LastX - m_StartX))
    End If
    
    ' Update the point.
     m_LastX = X
    If Shift = 2 Then
       m_LastY = m_StartY + Abs(m_LastX - m_StartX)
    Else
       m_LastY = Y
    End If

    ' Draw the new.
    m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
    If (m_LastX <> m_StartX) Then
       m_Canvas.Circle ((m_StartX + m_LastX) / 2, (m_StartY + m_LastY) / 2), varIIf(Abs(m_LastX - m_StartX) > Abs(m_LastY - m_StartY), Abs(m_LastX - m_StartX) / 2, Abs(m_LastY - m_StartY) / 2), , , , Abs((m_LastY - m_StartY) / (m_LastX - m_StartX))
    End If
    
End Sub

' Finish drawing the ellipse.
Private Sub m_Canvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
Dim cx As Single
Dim cy As Single
Dim rx As Single
Dim ry As Single
Dim ptx() As Single
Dim pty() As Single
Dim i As Integer
Dim dtheta As Single
Dim Pts() As PointAPI
    
    If Not m_DrawingStarted Then Exit Sub
    m_Canvas.AutoRedraw = True
    ' Erase the old.
    m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
    If (m_LastX <> m_StartX) Then
       m_Canvas.Circle ((m_StartX + m_LastX) / 2, (m_StartY + m_LastY) / 2), varIIf(Abs(m_LastX - m_StartX) > Abs(m_LastY - m_StartY), Abs(m_LastX - m_StartX) / 2, Abs(m_LastY - m_StartY) / 2), , , , Abs((m_LastY - m_StartY) / (m_LastX - m_StartX))
    End If
    
    ' Go back to vbCopyPen drawing mode.
    m_Canvas.DrawMode = vbCopyPen

    'Update point
    m_LastX = X
    If Shift = 2 Then
       m_LastY = m_StartY + Abs(m_LastX - m_StartX)
    Else
       m_LastY = Y
    End If
     m_StartX = m_StartX '/ gZoomFactor
    m_StartY = m_StartY '/ gZoomFactor
    m_LastX = m_LastX '/ gZoomFactor
    m_LastY = m_LastY '/ gZoomFactor
    Pts = EllipsePts(m_StartX, m_StartY, m_LastX - m_StartX, m_LastY - m_StartY, 0)
          
    ' Stop receiving events from the canvas.
    Set m_Canvas = Nothing

    ' Create the vbdPolygon that represents us.
    Set m_Polygon = New vbdDraw
    Set m_Object = m_Polygon
    
    With m_Polygon
         .NumPoints = 14 'm_MaxPoints
         
        For i = 0 To 12
           .X(i + 1) = Pts(i).X
           .Y(i + 1) = Pts(i).Y
           .TypePoint(i + 1) = 4
        Next
       
        .X(.NumPoints) = Pts(0).X
        .Y(.NumPoints) = Pts(0).Y
        .TypePoint(1) = 6
        .TypePoint(.NumPoints) = 3
    End With
      
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
            
    ' Me.vbdObject_ShadeX = m_ShadeX
    ' Me.vbdObject_ShadeY = m_ShadeY
    ' Me.vbdObject_ShadeFillStyle = m_ShadeFillStyle
    ' Me.vbdObject_ShadeGradient = m_ShadeGradient
    ' Me.vbdObject_ShadePattern = m_ShadePattern
          
    Me.vbdObject_MakeTransformation
           
    ' Tell the form to save us.
     AddObject Me
    
End Sub

' Add this transformation to the current one.
Public Sub vbdObject_AddTransformation(m() As Single)
    m_Object.AddTransformation m
    m_Object.MakeTransformation
End Sub

Private Property Let vbdObject_AlingText(ByVal RHS As Integer)

End Property

Private Property Get vbdObject_AlingText() As Integer

End Property

Private Property Let vbdObject_Angle(ByVal RHS As Single)
       m_Object.Angle = RHS
End Property

Private Property Get vbdObject_Angle() As Single
       vbdObject_Angle = m_Object.Angle
End Property

Public Property Let vbdObject_Bold(ByVal RHS As Boolean)
       
End Property

Public Property Get vbdObject_Bold() As Boolean

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
    m_Object.DrawInMetafile mf_dc
End Sub

Private Sub vbdObject_DrawPoint()

End Sub

Private Sub vbdObject_DrawSelect(ByVal canvas As PictureBox)
  m_Object.DrawSelect canvas
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

Private Function vbdObject_Info() As String
        vbdObject_Info = m_Object.Info
End Function

Public Property Get vbdObject_Italic() As Boolean

End Property

Public Property Let vbdObject_Italic(ByVal RHS As Boolean)

End Property

Public Sub vbdObject_MakeTransformation()
       m_Object.MakeTransformation
End Sub

Private Property Let vbdObject_MovePoints(ByVal RHS As Boolean)
        m_MovePoints = RHS
End Property

Private Property Get vbdObject_MovePoints() As Boolean
       vbdObject_MovePoints = m_MovePoints
End Property

Public Property Let vbdObject_Name(ByVal RHS As String)

End Property

Public Property Get vbdObject_Name() As String

End Property

Public Sub vbdObject_NewPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
           m_Object.NewPoint NumPoints, X(), Y(), TPoint()
End Sub

Public Sub vbdObject_NewTrPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
   Dim n As Long
      m_NumPoints = NumPoints
      ReDim m_OriginalPoints(1 To NumPoints)
      ReDim m_TypePoint(1 To NumPoints)
      For n = 1 To m_NumPoints
          m_OriginalPoints(n).X = X(n)
          m_OriginalPoints(n).Y = Y(n)
          m_TypePoint(n) = TPoint(n)
      Next
       m_Polygon.Identity
End Sub

Public Property Let vbdObject_ObjLock(ByVal RHS As Boolean)
             m_Object.ObjLock = RHS
End Property

Public Property Get vbdObject_ObjLock() As Boolean
          vbdObject_ObjLock = m_Object.ObjLock
End Property

Public Property Get vbdObject_Blend() As Integer
          vbdObject_Blend = m_Object.Blend
End Property

Public Property Let vbdObject_Blend(ByVal RHS As Integer)
          m_Object.Blend = RHS
End Property

Public Property Let vbdObject_Pattern(ByVal RHS As String)
          m_Object.Pattern = RHS
End Property

Public Property Get vbdObject_Pattern() As String
          vbdObject_Pattern = m_Object.Pattern
End Property

Public Property Set vbdObject_Picture(ByVal RHS As stdole.Picture)

End Property

Public Property Get vbdObject_Picture() As stdole.Picture

End Property

Public Sub vbdObject_ReadPoint(ByRef NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
       m_Object.ReadPoint NumPoints, X(), Y(), TPoint()
End Sub

Public Sub vbdObject_ReadTrPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
        m_Object.ReadTrPoint NumPoints, X(), Y(), TPoint()
End Sub

' Set the object's Selected status.
Private Property Let vbdObject_Selected(ByVal RHS As Boolean)
    m_Object.Selected = RHS
End Property

' Return the object's Selected status.
Private Property Get vbdObject_Selected() As Boolean
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
            Case "Polygon", "Picture", "RectAngle", "Polyline", "PolyDraw", "FreePolyline", "Scribble", "Scene", "Ellipse", "Text", "TextFrame", "TextPath"
                m_Object.Serialization = token_value
        End Select
    Loop
End Property
' Return a serialization string for the object.
Public Property Get vbdObject_Serialization() As String
    Dim txt As String
    txt = txt & "  IsBox(" & Format$(IsBox) & ") "
    txt = txt & m_Object.Serialization
    vbdObject_Serialization = "Ellipse(" & txt & ")"
End Property

Public Sub MaxPoint(Optional nMaxPoint As Integer = 10)
      m_MaxPoints = nMaxPoint
End Sub

Public Property Let vbdObject_Shade(ByVal RHS As Boolean)
     m_Shade = RHS
End Property

Public Property Get vbdObject_Shade() As Boolean
      vbdObject_Shade = m_Shade
End Property

Public Property Get vbdObject_Size() As Integer

End Property

Public Property Let vbdObject_Size(ByVal RHS As Integer)

End Property

Public Sub vbdObject_StartBound(xmin As Single, ymin As Single)
     m_Object.StartBound xmin, ymin
End Sub

Public Property Let vbdObject_Strikethrough(ByVal RHS As Boolean)

End Property

Public Property Get vbdObject_Strikethrough() As Boolean

End Property

Public Property Let vbdObject_TextDraw(ByVal RHS As String)

End Property

Public Property Get vbdObject_TextDraw() As String

End Property

Public Property Let vbdObject_TypeDraw(ByVal RHS As DrawType)
        m_Object.TypeDraw = RHS
End Property

Public Property Get vbdObject_TypeDraw() As DrawType
       vbdObject_TypeDraw = m_Object.TypeDraw
End Property

' Purpose    : Return varTrue if blnCondition = true, or varFalse otherwise
' Assumptions: -
' Effects    : -
' Inputs     : blnCondition, varTrue, varFalse
' Returns    : As specified
Public Function varIIf(blnCondition As Boolean, varTrue As Variant, varFalse As Variant) As Variant
  On Error GoTo ErrorHandler
  
  If blnCondition Then
    varIIf = varTrue
  Else
    varIIf = varFalse
  End If
  Exit Function

ErrorHandler:
     'ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Function

Private Function EllipsePts(cLeft As Single, cTop As Single, cWidth As Single, cHeight As Single, cAngle As Single, Optional MaxPoint As Integer = 12) As PointAPI()
    Dim offsetX As Single
    Dim offsetY As Single
    Dim r As Single
    Dim Alfa As Single
    Dim pX(12) As Single
    Dim pY(12) As Single
    Dim POINT(12) As PointAPI
    Dim n As Integer
    Dim CenterX As Single
    Dim CenterY As Single
    Dim eFactor As Double

    eFactor = 2 / 3 * (Sqr(2) - 1)

    offsetX = cWidth * eFactor
    offsetY = cHeight * eFactor
    CenterX = cWidth / 2
    CenterY = cHeight / 2
    
    pX(0) = cWidth
    pX(1) = pX(0)
    pX(11) = pX(0)
    pX(12) = pX(0)
    
    pX(5) = 0
    pX(6) = pX(5)
    pX(7) = pX(5)
    
    pX(2) = CenterX + offsetX
    pX(10) = pX(2)
    
    pX(4) = CenterX - offsetX
    pX(8) = pX(4)

    pX(3) = CenterX
    pX(9) = pX(3)
    
    pY(2) = 0
    pY(3) = pY(2)
    pY(4) = pY(2)
    
    pY(8) = cHeight
    pY(9) = pY(8)
    pY(10) = pY(8)
    
    pY(7) = CenterY + offsetY
    pY(11) = pY(7)
    
    pY(1) = CenterY - offsetY
    pY(5) = pY(1)
    
    pY(0) = CenterY
    pY(12) = pY(0)
    pY(6) = pY(0)
   
    
    For n = 0 To MaxPoint
        r = Sqr(pX(n) ^ 2 + pY(n) ^ 2)
        Alfa = m2Atn2(pY(n), pX(n)) - (cAngle * PI / 180)
        POINT(n).X = cLeft + r * Cos(Alfa)
        POINT(n).Y = cTop + r * Sin(Alfa)
    Next n
  '  POINT(13) = POINT(0)
   ' POINT(13) = POINT(0)
    EllipsePts = POINT
End Function

'
Private Sub DrawEllipse(cLeft As Single, cTop As Single, cWidth As Single, cHeight As Single, Optional cAngle As Single)
   Dim Pts() As PointAPI
   Dim OldScale As typScaleMode
  
   OldScale = SetWordScale(m_Canvas)

   Pts = EllipsePts(cLeft, cTop, cWidth, cHeight, cAngle)
   
   'BeginPath m_Canvas.hDC
   PolyBezier m_Canvas.hDC, Pts(0), UBound(Pts) + 1
   'EndPath m_Canvas.hDC
   'StrokeAndFillPath m_Canvas.hDC
    ResetWordScale m_Canvas, OldScale

End Sub

Public Property Let vbdObject_TypeFill(ByVal RHS As DrawTypeFill)
        m_TypeFill = RHS
End Property

Public Property Get vbdObject_TypeFill() As DrawTypeFill
        vbdObject_TypeFill = m_TypeFill
End Property

Public Property Let vbdObject_Underline(ByVal RHS As Boolean)

End Property

Public Property Get vbdObject_Underline() As Boolean

End Property

Public Property Let vbdObject_Weight(ByVal RHS As Integer)

End Property

Public Property Get vbdObject_Weight() As Integer

End Property


