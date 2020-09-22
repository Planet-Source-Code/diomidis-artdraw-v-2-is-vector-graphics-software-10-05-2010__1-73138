VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VbText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' VbDraw Line/Rectangle object.

Implements vbdObject

' Indicates a box rather than a line.
Public IsBox As Boolean

' The surface on which the user is clicking
' to define the object. This is set only during
' creation of this object.
Public WithEvents m_Canvas As PictureBox
Attribute m_Canvas.VB_VarHelpID = -1
Private m_DrawingStarted As Boolean

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
Public m_Blend As Integer
'Public m_Font As StdFont
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
Public m_Selected As Boolean

' Rubberband variables.
Private m_StartX As Single
Private m_StartY As Single
Private m_LastX As Single
Private m_LastY As Single


' Start drawing a rubberband box.
Private Sub m_Canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xmin As Single, ymin As Single, xmax As Single, ymax As Single
Dim iCounter As Long, i As Long
Dim r As Single
Dim Alfa As Single
Dim nfonts As New StdFont
Dim PointCoods() As PointAPI
Dim PointType() As Byte
'Dim xmin As Long, ymin As Long

   m_DrawingStarted = True

  If m_TypeDraw = dTextFrame Then
    ' Start using dotted vbInvert mode.
    m_Canvas.DrawMode = vbInvert
    m_Canvas.DrawStyle = vbDot

    ' Start the first rubberband box.
    m_StartX = X
    m_StartY = Y
    m_LastX = X
    m_LastY = Y
    If IsBox Then
        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
    Else
        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
    End If
    
  ElseIf m_TypeDraw = dText Then
    m_DrawingStarted = False
    m_StartX = m_StartX '/ gZoomFactor
    m_StartY = m_StartY '/ gZoomFactor
    m_LastX = m_LastX '/ gZoomFactor
    m_LastY = m_LastY '/ gZoomFactor

    ' Go back to vbCopyPen drawing mode.
    m_Canvas.DrawMode = vbCopyPen

    ' Create the vbdPolygon that represents us.
    Set m_Polygon = New vbdDraw
    Set m_Object = m_Polygon
    m_StartX = X
    m_StartY = Y
    m_CurrentX = m_StartX
    m_CurrentY = m_StartY

    IsBox = True

    nfonts.Bold = False
    nfonts.Charset = 0
    nfonts.Italic = False
    nfonts.Name = "Arial"
    nfonts.Size = 20
    nfonts.Strikethrough = False
    nfonts.Underline = False
    'nfonts.Weight = 400
    If FrmFonts.ShowForm(nfonts, m_TextDraw, m_AlingText) = False Then
       m_Bold = nfonts.Bold
       m_Charset = nfonts.Charset
       m_Italic = nfonts.Italic
       m_Name = nfonts.Name
       m_Size = nfonts.Size
       m_Strikethrough = nfonts.Strikethrough
       m_Underline = nfonts.Underline
       m_Weight = nfonts.Weight
              
       ' Create the vbdPolygon that represents us.
       Set m_Polygon = New vbdDraw
       Set m_Object = m_Polygon

       With m_Polygon
            .NumPoints = 2
            .X(1) = m_StartX
            .X(2) = m_LastX
            .Y(1) = m_StartY
            .Y(2) = m_StartY
            .IsClosed = True
            .TypePoint(1) = 6
            .TypePoint(2) = 3
        End With
    
        With m_Canvas
            .Font.Bold = nfonts.Bold
            .Font.Charset = nfonts.Charset
            .Font.Italic = nfonts.Italic
            .Font.Name = nfonts.Name
            .Font.Size = nfonts.Size
            .Font.Strikethrough = nfonts.Strikethrough
            .Font.Underline = nfonts.Underline
            .Font.Weight = nfonts.Weight
            .CurrentX = m_StartX
            .CurrentY = m_StartY
             m_CurrentX = m_StartX
             m_CurrentY = m_StartY
         End With
         xmax = m_Canvas.TextWidth(m_TextDraw)
         ymax = m_Canvas.TextHeight(m_TextDraw)
'         Stop

''         DrawChar m_Canvas, m_TextDraw, m_CurrentX, m_CurrentY, PointCoods, PointType, xmin, ymin
''          iCounter = UBound(PointCoods) '+ 1
''           With m_Polygon
''             .IsClosed = True
''             .NumPoints = iCounter
''             For i = 1 To iCounter
''                .X(i) = PointCoods(i).X - (xmin - m_StartX)
''                .Y(i) = PointCoods(i).Y - (ymin - m_StartY)
''                .TypePoint(i) = PointType(i)
''             Next
''         End With
          
        ReadPathText m_Canvas, m_TextDraw, PointCoods(), PointType(), iCounter
        
        'Find difference from start point
         With PointCoods(0)
            xmin = .X
            ymin = .Y
        End With
        For i = 1 To iCounter - 1
            With PointCoods(i)
                If xmin > .X Then xmin = .X
                If ymin > .Y Then ymin = .Y
            End With
        Next i
        
        'Stop
        
        With m_Polygon
            .IsClosed = True
            .NumPoints = iCounter
             For i = 1 To iCounter
                .X(i) = PointCoods(i - 1).X - (xmin - m_StartX)
                .Y(i) = PointCoods(i - 1).Y - (ymin - m_StartY)
                .TypePoint(i) = PointType(i - 1)
             Next
        End With
          
        
        
         
'         Stop
         'BoundText PointCoods, xmin, ymin, xMax, yMax
'         With m_Polygon
'             .NumPoints = 4
'             .X(1) = X - (xMax / 2)
'             .X(2) = X + (xMax / 2)
'             .X(3) = X + (xMax / 2)
'             .X(4) = X - (xMax / 2)
'             .Y(1) = Y - (yMax / 2)
'             .Y(2) = Y - (yMax / 2)
'             .Y(3) = Y + (yMax / 2)
'             .Y(4) = Y + (yMax / 2)
'             .IsClosed = True
'             .TypePoint(1) = 6
'             .TypePoint(2) = 2
'             .TypePoint(3) = 2
'             .TypePoint(4) = 3
'          End With
          Set m_Canvas = Nothing
    Else
       Set m_Canvas = Nothing
       Exit Sub
    End If

     Me.vbdObject_DrawStyle = m_DrawStyle
     Me.vbdObject_DrawWidth = m_DrawWidth
     Me.vbdObject_FillColor = RGB(255, 255, 255) 'm_FillColor
     Me.vbdObject_FillStyle = m_FillStyle
     Me.vbdObject_ForeColor = RGB(0, 0, 0) 'm_ForeColor
     Me.vbdObject_TextDraw = m_TextDraw
     Me.vbdObject_AlingText = m_AlingText
     Me.vbdObject_TypeDraw = m_TypeDraw
     Me.vbdObject_TypeFill = m_TypeFill
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
    ' Tell the form to save us.
    AddObject Me
 End If
    
End Sub

' Continue drawing the rubberband box.
Private Sub m_Canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_DrawingStarted Then Exit Sub

    ' Erase the old box.
    If IsBox Then
        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
    Else
        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
    End If

    ' Update the point.
       m_LastX = X
    If Shift = 2 Then
       m_LastY = m_StartY + Abs(m_LastX - m_StartX)
    Else
       m_LastY = Y
    End If
    
    ' Draw the new box.
    If IsBox Then
        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
    Else
        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
    End If
End Sub

' Finish drawing the rubberband box.
Private Sub m_Canvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim n As Integer
Dim iCounter As Long
Dim r As Single
Dim Alfa As Single
Dim nfonts As New StdFont
Dim PointCoods() As PointAPI
Dim PointType() As Byte

    If Not m_DrawingStarted Then Exit Sub
    
    ' Erase the old box.
    If IsBox Then
        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY), , B
    Else
        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
    End If
    
    ' Go back to vbCopyPen drawing mode.
    m_Canvas.DrawMode = vbCopyPen

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
    
    ' Go back to vbCopyPen drawing mode.
    m_Canvas.DrawMode = vbCopyPen
    ' Stop receiving events from the canvas.
    Set m_Canvas = Nothing

    nfonts.Bold = False
    nfonts.Charset = 0
    nfonts.Italic = False
    nfonts.Name = "Arial"
    nfonts.Size = 20
    nfonts.Strikethrough = False
    nfonts.Underline = False
    nfonts.Weight = 400

    If FrmFonts.ShowForm(nfonts, m_TextDraw, m_AlingText) = False Then
        m_Bold = nfonts.Bold
        m_Charset = nfonts.Charset
        m_Italic = nfonts.Italic
        m_Name = nfonts.Name
        m_Size = nfonts.Size
        m_Strikethrough = nfonts.Strikethrough
        m_Underline = nfonts.Underline
        m_Weight = nfonts.Weight
        
        ' Create the vbdPolygon that represents us.
        Set m_Polygon = New vbdDraw
        Set m_Object = m_Polygon

             With m_Polygon
                .NumPoints = 4
                .X(1) = m_StartX
                .X(2) = m_LastX
                .X(3) = m_LastX
                .X(4) = m_StartX
                .Y(1) = m_StartY
                .Y(2) = m_StartY
                .Y(3) = m_LastY
                .Y(4) = m_LastY
                .IsClosed = True
                .TypePoint(1) = 6
                .TypePoint(2) = 2
                .TypePoint(3) = 2
                .TypePoint(4) = 3
            End With
     Else
      
        Exit Sub
    End If
    
    Me.vbdObject_DrawStyle = m_DrawStyle
    Me.vbdObject_DrawWidth = m_DrawWidth
    Me.vbdObject_FillColor = RGB(255, 255, 255) 'm_FillColor
    Me.vbdObject_FillColor2 = RGB(155, 155, 155)
    Me.vbdObject_FillStyle = m_FillStyle
    Me.vbdObject_ForeColor = RGB(0, 0, 0) 'm_ForeColor
    Me.vbdObject_TextDraw = m_TextDraw
    Me.vbdObject_AlingText = m_AlingText
    Me.vbdObject_TypeDraw = m_TypeDraw
    Me.vbdObject_TypeFill = m_TypeFill
    Me.vbdObject_Shade = m_Shade
    Me.vbdObject_Blend = m_Blend
    Me.vbdObject_Bold = m_Bold
    Me.vbdObject_Charset = m_Charset
    Me.vbdObject_Italic = m_Italic
    Me.vbdObject_Name = m_Name
    Me.vbdObject_Size = m_Size
    Me.vbdObject_Strikethrough = m_Strikethrough
    Me.vbdObject_Underline = m_Underline
    Me.vbdObject_Weight = m_Weight
    Me.vbdObject_MakeTransformation
    
    ' Tell the form to save us.
    AddObject Me

End Sub


' Add this transformation to the current one.
Private Sub vbdObject_AddTransformation(m() As Single)
    m_Object.AddTransformation m
    m_Object.MakeTransformation
End Sub

Public Property Let vbdObject_AlingText(ByVal RHS As Integer)
        m_Object.AlingText = RHS
End Property

Public Property Get vbdObject_AlingText() As Integer
        vbdObject_AlingText = m_Object.AlingText
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
        m_Charset = RHS
End Property

Public Property Get vbdObject_Charset() As Integer
         vbdObject_Charset = m_Charset
End Property

' Clear the object's transformation.
Private Sub vbdObject_ClearTransformation()
    m_Object.ClearTransformation
End Sub

Private Property Let vbdObject_CurrentX(ByVal RHS As Single)
          m_CurrentX = RHS
End Property

Private Property Get vbdObject_CurrentX() As Single
            vbdObject_CurrentX = m_CurrentX
End Property

Private Property Let vbdObject_CurrentY(ByVal RHS As Single)
          m_CurrentY = RHS
End Property

Private Property Get vbdObject_CurrentY() As Single
         vbdObject_CurrentY = m_CurrentY
End Property

' Draw the object in a metafile.
Private Sub vbdObject_DrawInMetafile(ByVal mf_dc As Long)
    m_Object.DrawInMetafile mf_dc
End Sub

Private Sub vbdObject_DrawPoint()

End Sub

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

Public Property Let vbdObject_FillColor2(ByVal RHS As OLE_COLOR)
        m_Object.FillColor2 = RHS
End Property

Public Property Get vbdObject_FillColor2() As OLE_COLOR
        vbdObject_FillColor2 = m_Object.FillColor2
End Property

Private Property Let vbdObject_FillMode(ByVal RHS As FillMode)
         m_Object.FillMode = RHS
End Property

Private Property Get vbdObject_FillMode() As FillMode
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

Private Property Let vbdObject_Gradient(ByVal RHS As Integer)
        m_Object.Gradient = RHS
End Property

Private Property Get vbdObject_Gradient() As Integer
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

Public Property Let vbdObject_Italic(ByVal RHS As Boolean)
        m_Object.Italic = RHS
End Property

Public Property Get vbdObject_Italic() As Boolean
          vbdObject_Italic = m_Object.Italic
End Property

Public Sub vbdObject_MakeTransformation()
        m_Object.MakeTransformation
End Sub

Public Property Let vbdObject_Name(ByVal RHS As String)
        m_Object.Name = RHS
End Property

Public Property Get vbdObject_Name() As String
          vbdObject_Name = m_Object.Name
End Property

Private Sub vbdObject_NewPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)

             m_Object.NewPoint NumPoints, X(), Y(), TPoint()

'              Dim n As Long
'              Dim StartX As Single
'              Dim StartY As Single
'              m_Polygon.NumPoints = NumPoints
'
'              For n = 1 To NumPoints
'                 m_Polygon.X(n) = X(n)
'                 m_Polygon.Y(n) = Y(n)
'                 m_Polygon.TypePoint(n) = TPoint(n - 1)
'              Next
'               m_Polygon.Identity
End Sub

Private Sub vbdObject_NewTrPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
' Dim n As Long
'      m_NumPoints = NumPoints
'
'      For n = 1 To m_NumPoints
'          m_OriginalPoints(n).X = X(n)
'          m_OriginalPoints(n).Y = Y(n)
'          m_TransformedPoints(n) = TypePoint(n)
'      Next
'     m_Polygon.Identity
End Sub

Private Property Let vbdObject_ObjLock(ByVal RHS As Boolean)
            m_Object.ObjLock = RHS
End Property

Private Property Get vbdObject_ObjLock() As Boolean
            vbdObject_ObjLock = m_Object.ObjLock
End Property

Public Property Let vbdObject_Blend(ByVal RHS As Integer)
            m_Object.Blend = RHS
End Property

Public Property Get vbdObject_Blend() As Integer
            vbdObject_Blend = m_Object.Blend
End Property

Private Property Let vbdObject_Pattern(ByVal RHS As String)
       m_Object.Pattern = RHS
End Property

Private Property Get vbdObject_Pattern() As String
         vbdObject_Pattern = m_Object.Pattern
End Property

Private Property Set vbdObject_Picture(ByVal RHS As stdole.Picture)

End Property

Private Property Get vbdObject_Picture() As stdole.Picture

End Property

Private Sub vbdObject_ReadPoint(ByRef NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)

End Sub

Private Sub vbdObject_ReadTrPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
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
            Case "Polyline", "Scribble", "FreePolygon", "PolyDraw", "Polygon", "Picture", "RectAngle", "Ellipse", "Text", "TextFrame", "TextPath", "Scene"
                m_Object.Serialization = token_value
        End Select
    Loop
End Property

' Return a serialization string for the object.
Public Property Get vbdObject_Serialization() As String
Dim txt As String

    txt = txt & "  IsBox(" & Format$(IsBox) & ") "
    txt = txt & m_Object.Serialization
    Select Case m_Object.TypeDraw
    Case dText
       vbdObject_Serialization = "Text(" & txt & ")"
    Case dTextFrame
       vbdObject_Serialization = "TextFrame(" & txt & ")"
'    Case dPolydraw
'       vbdObject_Serialization = "PolyDraw(" & txt & ")"
'    Case dPolyline
'       vbdObject_Serialization = "Polyline(" & txt & ")"
    Case Else
'       Stop
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

Private Sub vbdObject_StartBound(xmin As Single, ymin As Single)
      m_Object.StartBound xmin, ymin
End Sub

Public Property Get vbdObject_Strikethrough() As Boolean
         vbdObject_Strikethrough = m_Object.Strikethrough
End Property

Public Property Let vbdObject_Strikethrough(ByVal RHS As Boolean)
        m_Object.Strikethrough = RHS
End Property

Public Property Let vbdObject_TextDraw(ByVal RHS As String)
      m_Object.TextDraw = RHS
End Property

Public Property Get vbdObject_TextDraw() As String
     vbdObject_TextDraw = m_Object.TextDraw
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
        
End Property

Public Property Get vbdObject_TypeFill() As DrawTypeFill
'
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


'Read Path text and make PointCoolds and Type for draw
Private Sub ReadPathText(Obj As PictureBox, _
                         txt As String, _
                         Point_Coords() As PointAPI, _
                         Point_Types() As Byte, _
                         NumPoints As Long)
    Dim ret As Long
    ret = BeginPath(Obj.hDC)
    Obj.Print txt
    ret = EndPath(Obj.hDC)
    NumPoints = 0
    NumPoints = GetPathAPI(Obj.hDC, ByVal 0&, ByVal 0&, 0)

    If (NumPoints) Then
        ReDim Point_Coords(NumPoints - 1)
        ReDim Point_Types(NumPoints - 1)
        'Get the path data from the DC
        Call GetPathAPI(Obj.hDC, Point_Coords(0), Point_Types(0), NumPoints)
    End If

End Sub

