VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdScribble"
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
Public m_DrawingEdit As Boolean
Private m_DrawingMove As Boolean
Private m_SelectPoint As Integer

' Constituent vbdPolygon object.
Private m_Polygon As vbdDraw
Private m_Object As vbdObject

Public IsClosed As Boolean

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
Private m_OriginalPoints2() As PointAPI 'calligraphic
Private m_TransformedPoints() As PointAPI
Private m_Lenght As Single 'calligraphic
Private i As Integer

' Rubberband variables.
Private m_StartX As Single
Private m_StartY As Single
Private m_LastX As Single
Private m_LastY As Single

Private Sub OptimizePoint()
      Dim aa As Integer, n As Integer, m_NewPoints() As PointAPI, stmp As PointAPI, etmp As PointAPI
      Dim pv As Integer
st1:
      aa = 0
       For n = 1 To m_NumPoints - 2 Step 2
      
          If n = 1 Then stmp = m_OriginalPoints(n): etmp = m_OriginalPoints(n + 1)
          
          If stmp.X >= stmp.X - 5 And stmp.X <= stmp.X + 5 And _
             stmp.Y >= stmp.Y - 5 And etmp.Y <= etmp.Y + 5 Then
             aa = aa + 1
             
             If aa = 1 Then
                ReDim m_NewPoints(1 To aa)
                m_NewPoints(aa) = stmp
                stmp = etmp: etmp = m_OriginalPoints(n + 2)
             Else
                ReDim Preserve m_NewPoints(1 To aa)
                m_NewPoints(aa) = etmp
                stmp = etmp: etmp = m_OriginalPoints(n + 2)
             End If
         
          End If
      Next
      
      If aa Mod 3 = 0 Then
         aa = aa + 1
         ReDim Preserve m_NewPoints(1 To aa)
         m_NewPoints(aa) = m_NewPoints(aa - 1)
      ElseIf aa Mod 3 = 2 Then
         aa = aa - 1
         ReDim Preserve m_NewPoints(1 To aa)
      End If
      
      If m_TypeDraw = dCalligraphic Then
         aa = aa + 1
         ReDim Preserve m_NewPoints(1 To aa)
         m_NewPoints(aa) = m_NewPoints(1)
      End If
                     
      m_OriginalPoints = m_NewPoints
      
      ReDim m_TypePoint(1 To aa)
      m_TypePoint(1) = 6
      For n = 2 To aa
         m_TypePoint(n) = 4
      Next
      If m_TypeDraw = dCalligraphic Then
         m_TypePoint(aa) = 3
      End If
      m_NumPoints = aa
      DrawPoint
'      For n = 1 To m_NumPoints - 4 Step 2
'
'          If n = 1 Then stmp = m_OriginalPoints(n): etmp = m_OriginalPoints(n + 1)
'
'          If etmp.X >= stmp.X + 10 Or etmp.Y >= stmp.Y + 10 Then
'             aa = aa + 1
'
'             If aa = 1 Then
'                ReDim m_NewPoints(1 To aa)
'                m_NewPoints(aa) = stmp
'                stmp = etmp: etmp = m_OriginalPoints(n + 2)
'             Else
'                ReDim Preserve m_NewPoints(1 To aa)
'                m_NewPoints(aa) = etmp
'                stmp = etmp: etmp = m_OriginalPoints(n + 2)
'             End If
'          Else
'             etmp = m_OriginalPoints(n + 2)
'          End If
'      Next
'
'      If aa Mod 3 = 0 Then
'         aa = aa + 1
'         ReDim Preserve m_NewPoints(1 To aa)
'         m_NewPoints(aa) = m_OriginalPoints(m_NumPoints)
'      ElseIf aa Mod 3 = 2 Then
'         aa = aa - 1
'         ReDim Preserve m_NewPoints(1 To aa)
'      End If
'      m_OriginalPoints = m_NewPoints
'      ReDim m_TypePoint(1 To aa)
'      m_TypePoint(1) = 6
'      For n = 2 To aa
'         m_TypePoint(n) = 4
'      Next
'      m_NumPoints = aa
'      DrawPoint
          
End Sub

Private Sub AddPoint(NumPoint As Long, MX As Single, mY As Single)
      
      If NumPoint = 1 Then
         'Stop
         ReDim m_OriginalPoints(1 To NumPoint)
         If m_TypeDraw = dCalligraphic Then
            ReDim m_OriginalPoints2(1 To NumPoint)
            ReDim m_TypePoint(1 To NumPoint * 2)
         Else
            ReDim m_TypePoint(1 To NumPoint)
         End If
      Else
        ReDim Preserve m_OriginalPoints(1 To NumPoint)
        ReDim Preserve m_TypePoint(1 To NumPoint)
        If m_TypeDraw = dCalligraphic Then
            ReDim Preserve m_OriginalPoints2(1 To NumPoint)
            ReDim Preserve m_TypePoint(1 To NumPoint * 2)
        End If
      End If
      
      If m_TypeDraw = dScribble Then
         m_OriginalPoints(NumPoint).X = MX
         m_OriginalPoints(NumPoint).Y = mY
      Else
         If m_Lenght = 0 Then m_Lenght = 10
         Dim P As PointAPI
         'P = Rotate(mX - m_Lenght / 2, mY, 45)
         m_OriginalPoints(NumPoint).X = MX - m_Lenght / 2
         m_OriginalPoints(NumPoint).Y = mY
         'P = Rotate(mX + m_Lenght / 2, mY, 45)
         m_OriginalPoints2(NumPoint).X = MX + m_Lenght / 2
         m_OriginalPoints2(NumPoint).Y = mY '- (mY - m_Lenght) / 2
      End If
End Sub

Private Sub DrawPoint(Optional is_Draw As Boolean = False)
      
         Dim OldScale As typScaleMode, mStep As sinle
  
         OldScale = SetWordScale(m_Canvas)

         PolyDraw m_Canvas.hDC, m_OriginalPoints(1), m_TypePoint(1), m_NumPoints
      
        ResetWordScale m_Canvas, OldScale
        
        mStep = IIf(GAP / gZoomFactor > 0.5, GAP / gZoomFactor, 1)
              
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
           ' Me.m_ShadeX=Me.vbdObject_ShadeX
           ' Me.m_ShadeY=Me.vbdObject_ShadeY
           ' Me.m_ShadeFillStyle=Me.vbdObject_ShadeFillStyle
           ' Me.m_ShadeGradient=Me.vbdObject_ShadeGradient
           ' Me.m_ShadePattern=Me.vbdObject_ShadePattern
           ' Me.m_Gradient  = Me.vbdObject_Gradient
       End If
End Sub

Private Sub PolyPoints(nPoint As Integer, cx As Single, cy As Single)

    If nPoint > 0 Then
        m_OriginalPoints(nPoint).X = cx '/ gZoomFactor
        m_OriginalPoints(nPoint).Y = cy '/ gZoomFactor
    End If
    ''Debug.Print nPoint, m_OriginalPoints(nPoint).X, m_OriginalPoints(nPoint).Y
End Sub

Private Sub m_Canvas_KeyUp(KeyCode As Integer, Shift As Integer)
       Dim Points() As PointAPI, aa As Long
       ' Stop
       If KeyCode = vbKeyDelete Then
          If m_NumPoints > 3 Then
          If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
             m_NumPoints = m_NumPoints - 1
             ReDim Points(1 To m_NumPoints)
             aa = 0
             For i = 1 To m_NumPoints + 1
                If m_SelectPoint <> i Then
                  aa = aa + 1
                  Points(aa).X = m_OriginalPoints(i).X
                  Points(aa).Y = m_OriginalPoints(i).Y
                End If
             Next
             m_OriginalPoints = Points
             
          End If
          End If
          DrawPoint
       End If
       
End Sub

Private Sub m_Canvas_LostFocus()
Dim MX As Single, mY As Single
     
         Set m_Polygon = New vbdDraw
        Set m_Object = m_Polygon
        m_Polygon.IsClosed = IsClosed
       ' OptimizePoint
        With m_Polygon
            .NumPoints = m_NumPoints
            For i = 1 To m_NumPoints
                MX = m_OriginalPoints(i).X
                mY = m_OriginalPoints(i).Y
                PolyPoints i, MX, mY
                .X(i) = m_OriginalPoints(i).X
                .Y(i) = m_OriginalPoints(i).Y
                .TypePoint(i) = m_TypePoint(i)
            Next
        End With
        ' Stop receiving events from the canvas.
        Set m_Canvas = Nothing

         If m_TypeDraw = dPicture Then
         '  Me.vbdObject_Picture = m_Picture
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
     
     ' Me.vbdObject_ShadeX = m_ShadeX
     ' Me.vbdObject_ShadeY = m_ShadeY
     ' Me.vbdObject_ShadeFillStyle = m_ShadeFillStyle
     ' Me.vbdObject_ShadeGradient = m_ShadeGradient
     ' Me.vbdObject_ShadePattern = m_ShadePattern
          
     Me.vbdObject_MakeTransformation
        ' Tell the form to save us.
        AddObject Me
End Sub

' Start the scribble.
Private Sub m_Canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Start drawing the scribble.
    'm_DrawingStarted = True
    If m_DrawingEdit = False Then
        ' Create the vbdPolygon that represents us.
'        Set m_Polygon = New vbdDraw
'        Set m_Object = m_Polygon
'        m_Polygon.IsClosed = False

        ' Remember where we are.
        m_LastX = X '/ gZoomFactor
        m_LastY = Y '/ gZoomFactor
        
        m_NumPoints = 1
        AddPoint m_NumPoints, m_LastX, m_LastY
        ' Save this point.
'        m_Polygon.NumPoints = 1
'        m_Polygon.X(m_Polygon.NumPoints) = m_LastX
'        m_Polygon.Y(m_Polygon.NumPoints) = m_LastY
'        m_Polygon.TypePoint(m_Polygon.NumPoints) = 6
        ' Draw the line.
        'DrawPoint
        m_Canvas.CurrentX = X '/ gZoomFactor
        m_Canvas.CurrentY = Y '/ gZoomFactor
        m_DrawingStarted = True
       
    Else
         m_SelectPoint = 0
        For i = 1 To m_NumPoints
            If X >= m_OriginalPoints(i).X - 2 And X <= m_OriginalPoints(i).X + 2 And Y >= m_OriginalPoints(i).Y - 2 And Y <= m_OriginalPoints(i).Y + 2 Then
               m_SelectPoint = i
               Exit For
            End If
        Next
        
        If m_SelectPoint > 0 Then
           m_Canvas.DrawMode = vbInvert
           m_Canvas.DrawStyle = vbDot

'           m_LastX = X'/ gZoomFactor
'           m_LastY = Y'/ gZoomFactor
           
           DrawPoint
           m_DrawingMove = True
        Else
           'End Draw
           m_Canvas_LostFocus
        End If
    End If
    
End Sub

' Continue the scribble.
Private Sub m_Canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    'If Not m_DrawingStarted Then Exit Sub
    If m_DrawingStarted Then

        ' Make sure we really moved.
        If m_LastX = X And m_LastY = Y Then Exit Sub
    
        ' Remember where we are.
        m_LastX = X '/ gZoomFactor
        m_LastY = Y '/ gZoomFactor
        m_NumPoints = m_NumPoints + 1
        AddPoint m_NumPoints, m_LastX, m_LastY
        ' Save this point.
        'm_Polygon.NumPoints = m_Polygon.NumPoints + 1
        'm_Polygon.X(m_Polygon.NumPoints) = m_LastX 'X'/ gZoomFactor
        'm_Polygon.Y(m_Polygon.NumPoints) = m_LastY 'Y'/ gZoomFactor
        'm_Polygon.TypePoint(m_Polygon.NumPoints) = 2
       ' DrawPoint
        ' Draw the line.
        m_Canvas.Line -(m_LastX, m_LastY)
        
    ElseIf m_DrawingEdit Then
        If m_DrawingMove = False Then Exit Sub
                
        ' Start using dotted vbInvert mode.
      '  m_Canvas.DrawMode = vbInvert
        
        'Call PolyPoints(m_SelectPoint, m_LastX, m_LastY)
        DrawPoint
        ' Update the point.
        m_LastX = X '* gZoomFactor
        m_LastY = Y '* gZoomFactor

        Call PolyPoints(m_SelectPoint, m_LastX, m_LastY)
        DrawPoint
              
   End If
End Sub

' Finish the scribble.
Private Sub m_Canvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'If Not m_DrawingStarted Then Exit Sub
     If m_DrawingStarted Then
         If m_TypeDraw = dCalligraphic Then
              ReDim Preserve m_OriginalPoints(1 To m_NumPoints * 2 + 1)
              ReDim Preserve m_TypePoint(1 To m_NumPoints * 2 + 1)
              Dim aa As Long
              For i = m_NumPoints To 1 Step -1
                  aa = aa + 1
                  m_OriginalPoints(m_NumPoints + aa) = m_OriginalPoints2(i)
              Next
              m_NumPoints = m_NumPoints * 2 + 1
              m_OriginalPoints(m_NumPoints) = m_OriginalPoints(1)
         End If
         
          DrawPoint
           m_DrawingStarted = False
           m_DrawingEdit = True
           OptimizePoint
           DrawPoint
           m_Canvas_LostFocus
     
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
              m_LastY = Y '* gZoomFactor
          End If
          m_Canvas_Paint
          Call PolyPoints(m_SelectPoint, m_LastX, m_LastY)
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

' Add this transformation to the current one.
Private Sub vbdObject_AddTransformation(m() As Single)
    m_Object.AddTransformation m
    m_Object.MakeTransformation
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

Public Property Set vbdObject_Canvas(ByVal RHS As PictureBox)
    Set m_Canvas = RHS
End Property

Public Property Get vbdObject_Canvas() As PictureBox
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

Private Function vbdObject_Info() As String
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

Public Property Let vbdObject_Name(ByVal RHS As String)
          m_Object.Name = RHS
End Property

Public Property Get vbdObject_Name() As String
          vbdObject_Name = m_Object.Name
End Property

Public Sub vbdObject_NewPoint(NumPoints As Long, X() As Single, Y() As Single, TypePoint() As Byte)
        m_Object.NewPoint NumPoints, X(), Y(), TypePoint()

'           Dim n As Long
'              m_Polygon.NumPoints = NumPoints
'
'              For n = 1 To NumPoints
'                 m_Polygon.X(n) = X(n)'/ gZoomFactor '+ 'm_Object.CurrentX
'                 m_Polygon.Y(n) = Y(n)'/ gZoomFactor '+ m_Object.CurrentY
'                 m_Polygon.TypePoint(n) = TypePoint(n)
'              Next
'             m_Polygon.Identity
End Sub

Public Sub vbdObject_NewTrPoint(NumPoints As Long, X() As Single, Y() As Single, TypePoint() As Byte)
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

Public Property Set vbdObject_Picture(ByVal RHS As stdole.Picture)
        Set m_Object.Picture = RHS
End Property

Public Property Get vbdObject_Picture() As stdole.Picture
        Set vbdObject_Picture = m_Object.Picture
End Property

Public Sub vbdObject_ReadPoint(ByRef NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
        m_Object.ReadPoint NumPoints, X(), Y(), TPoint()
End Sub

Public Sub vbdObject_ReadTrPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
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
' The serialization does not include the ObjectType(...) part.
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
            Case "Polygon", "RectAngle", "Picture", "Polyline", "PolyDraw", "FreePolyline", "Scribble", "Scene", "Ellipse", "Text", "TextFrame", "TextPath"
                m_Object.Serialization = token_value
        End Select
    Loop
End Property
' Return a serialization string for the object.
Public Property Get vbdObject_Serialization() As String
Dim txt As String
    txt = txt & "  IsBox(" & Format$(IsBox) & ") "
    txt = txt & m_Object.Serialization
    vbdObject_Serialization = "Scribble(" & txt & ")"
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

Public Property Let vbdObject_Strikethrough(ByVal RHS As Boolean)
       m_Object.Strikethrough = RHS
End Property

Private Property Get vbdObject_Strikethrough() As Boolean
       vbdObject_Strikethrough = m_Object.Strikethrough
End Property

Public Property Let vbdObject_TextDraw(ByVal RHS As String)
       m_Object.TextDraw = RHS
End Property

Public Property Get vbdObject_TextDraw() As String
       vbdObject_TextDraw = m_Object.TextDraw
End Property

Public Property Let vbdObject_TypeDraw(ByVal RHS As DrawType)
        m_Object.TypeDraw = RHS
End Property

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

