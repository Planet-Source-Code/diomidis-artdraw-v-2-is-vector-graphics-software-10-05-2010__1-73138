VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdScene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' VbDraw scene object.

Implements vbdObject

' The objects in the scene.
Public SceneObjects As Collection

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
Public m_hRegion As Long
Public m_Picture As StdPicture
Public m_info As String
' Move these objects to the beginning of the
' SceneObjects collection so they are drawn
' first.
Public Sub MoveToBack(ByVal targets As Collection)
Dim target As vbdObject

    ' Remove the objects from SceneObjects.
    RemoveObjects targets

    ' Re-add the objects at the beginning.
    For Each target In targets
        If SceneObjects.Count < 1 Then
            SceneObjects.Add target
        Else
            SceneObjects.Add target, , 1
        End If
    Next target
End Sub

' Move these objects to the beginning of the
' SceneObjects collection so they are drawn Backward
Public Sub MoveToBackward(ByVal targets As Collection)
Dim target As vbdObject
Dim newtarget As Collection
Dim Obj As vbdObject
Dim i As Integer
    
    ' Find the objects from SceneObjects.
    For Each target In targets
        ' Find this target.
        i = 1
        For Each Obj In SceneObjects
            If Obj Is target Then
               SceneObjects.Remove i
                'Exit For
               GoTo mtb
            End If
            i = i + 1
        Next Obj
    Next target
mtb:
    If i - 1 <= 0 Then i = 1 Else i = i - 1
    ' Re-add the objects at the beginning.
    For Each target In targets
        If SceneObjects.Count < 1 Then
           SceneObjects.Add target
        Else
           SceneObjects.Add target, , i
        End If
    Next target

End Sub

' Move these objects to the beginning of the
' SceneObjects collection so they are drawn Foward
Public Sub MoveToFoward(ByVal targets As Collection)
Dim target As vbdObject
Dim newtarget As Collection
Dim Obj As vbdObject
Dim i As Integer
    
    ' Find the objects from SceneObjects.
    For Each target In targets
        ' Find this target.
        i = 1
        For Each Obj In SceneObjects
            If Obj Is target Then
               If i = SceneObjects.Count Then Exit Sub
               SceneObjects.Remove i
               GoTo Mtf:
            End If
            i = i + 1
        Next Obj
    Next target
Mtf:
     i = i + 1
     'If i > SceneObjects.Count Then i = SceneObjects.Count
     ' Re-add the objects at the beginning.
      For Each target In targets
          If SceneObjects.Count < i Then
             SceneObjects.Add target
          Else
             SceneObjects.Add target, , i
          End If
      Next target
    
End Sub

' Paste objects of theSceneObjects collection so they are drawn last.
Public Sub PasteObject(ByVal targets As Collection)
    Dim target As vbdObject

    ' Add the objects .
    For Each target In targets
        SceneObjects.Add target
    Next target
End Sub

' Move these objects to the end of the
' SceneObjects collection so they are drawn
' last.
Public Sub MoveToFront(ByVal targets As Collection)
    Dim target As vbdObject

    ' Remove the objects from SceneObjects.
    RemoveObjects targets

    ' Re-add the objects at the end.
    For Each target In targets
        SceneObjects.Add target
    Next target
End Sub

' Remove these objects from SceneObjects.
Public Sub RemoveObjects(ByVal targets As Collection)
Dim target As vbdObject
Dim Obj As vbdObject
Dim i As Integer, F As Integer

    ' Remove the objects from SceneObjects.
    For Each target In targets
        ' Find this target.
        i = 1
        For Each Obj In SceneObjects
            If Obj Is target Then
               'If Obj.ObjLock = False Then
                 ' f = MsgBox("Delete object?", vbQuestion + vbYesNo + vbDefaultButton2)
                 '   If f = 6 Then
                      SceneObjects.Remove i
               '    End If
                    'Exit For
               'Else
              '    MsgBox "Object is Lock.", vbExclamation
              ' End If
            End If
            i = i + 1
        Next Obj
    Next target
End Sub

'Select object
Public Function SelectObj() As vbdObject
Dim Obj As vbdObject
Dim i As Integer
    'SelectObj = 0
    Set SelectObj = Nothing

    ' Search for the object starting with
    ' the objects on top.
    For i = SceneObjects.Count To 1 Step -1
        Set Obj = SceneObjects(i)
        If Obj.Selected Then
            Set SelectObj = Obj
            Exit For
        End If
    Next i
End Function


'Select the object .
Public Sub SelectAllObject()
Dim Obj As vbdObject
Dim i As Integer

    ' Search for the object starting with the objects on top.
    For i = SceneObjects.Count To 1 Step -1
        Set Obj = SceneObjects(i)
        Obj.Selected = False
        SelectVbdObject Obj
    Next i
End Sub

' Deselect the object.
Public Sub DeselectAllObject()
Dim Obj As vbdObject
Dim i As Integer

    ' Search for the object starting with the objects on top.
    For i = SceneObjects.Count To 1 Step -1
        Set Obj = SceneObjects(i)
        Obj.Selected = False
        Obj.EditPoint = False
        DeselectVbdObject Obj
    Next i
End Sub

Public Sub NewTransformation()
     Dim Obj As vbdObject
     Dim i As Integer

    ' Search for the object starting with the objects on top.
    For i = SceneObjects.Count To 1 Step -1
        Set Obj = SceneObjects(i)
        Obj.MakeTransformation
    Next i

End Sub

' Add this transformation to the current one.
Private Sub vbdObject_AddTransformation(m() As Single)
Dim Obj As vbdObject

    For Each Obj In SceneObjects
        Obj.AddTransformation m
    Next Obj
End Sub

Private Property Let vbdObject_AlingText(ByVal RHS As Integer)

End Property

Private Property Get vbdObject_AlingText() As Integer

End Property

Private Property Let vbdObject_Angle(ByVal RHS As Single)

End Property

Private Property Get vbdObject_Angle() As Single

End Property

Public Property Let vbdObject_Bold(ByVal RHS As Boolean)
         m_Bold = RHS
End Property

Public Property Get vbdObject_Bold() As Boolean
         vbdObject_Bold = m_Bold
End Property

Private Property Set vbdObject_Canvas(ByVal RHS As PictureBox)
    ' Do nothing. This object is not directly
    ' creatable by the user.
End Property

Private Property Get vbdObject_Canvas() As PictureBox
    ' Do nothing. This object is not directly
    ' creatable by the user.
    Set vbdObject_Canvas = Nothing
End Property

Public Property Let vbdObject_Charset(ByVal RHS As Integer)
          m_Charset = RHS
End Property

Public Property Get vbdObject_Charset() As Integer
         vbdObject_Charset = m_Charset
End Property

' Clear the object's transformation.
Private Sub vbdObject_ClearTransformation()
   Dim Obj As vbdObject
    For Each Obj In SceneObjects
        Obj.ClearTransformation
    Next Obj
End Sub

Private Property Get vbdObject_CurrentX() As Single

End Property

Private Property Let vbdObject_CurrentX(ByVal RHS As Single)

End Property

Private Property Get vbdObject_CurrentY() As Single

End Property

Private Property Let vbdObject_CurrentY(ByVal RHS As Single)

End Property

' Draw the object in a metafile.
Private Sub vbdObject_DrawInMetafile(ByVal mf_dc As Long)
Dim Obj As vbdObject

    For Each Obj In SceneObjects
        Obj.DrawInMetafile mf_dc
    Next Obj
End Sub

Private Sub Class_Initialize()
    Set SceneObjects = New Collection
End Sub

' Return this object's bounds.
Private Sub vbdObject_Bound(xmin As Single, ymin As Single, xmax As Single, ymax As Single)
    BoundObjects SceneObjects, xmin, ymin, xmax, ymax
End Sub

Private Sub vbdObject_DrawPoint()

End Sub

Private Sub vbdObject_DrawSelect(ByVal canvas As PictureBox)
   
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
          ' m_EditPoint = RHS
End Property

Private Property Get vbdObject_EditPoint() As Boolean
          'vbdObject_EditPoint = m_EditPoint
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

' Draw the object on the canvas.
Private Sub vbdObject_Draw(ByVal canvas As PictureBox, Optional SetWsc As Boolean = True)
   Dim Obj As vbdObject
   Dim OldScale As typScaleMode
   If SetWsc Then
   OldScale = SetWordScale(canvas)
   End If
   
   For Each Obj In SceneObjects
       Obj.Draw canvas
   Next Obj
   
   For Each Obj In SceneObjects
      If Obj.Selected Or Obj.EditPoint Then
          If Obj.EditPoint Then
             ResetWordScale canvas, OldScale
          End If
          Obj.DrawSelect canvas
      End If
   Next Obj
   
'   ClearDC canvas.hdc
   If SetWsc Then
      ResetWordScale canvas, OldScale
    End If
  
End Sub

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
        vbdObject_Info = m_info
End Function

Public Property Let vbdObject_Italic(ByVal RHS As Boolean)
         m_Italic = RHS
End Property

Public Property Get vbdObject_Italic() As Boolean
          vbdObject_Italic = m_Italic
End Property

Public Sub vbdObject_MakeTransformation()
     Dim Obj As vbdObject
     For Each Obj In SceneObjects
      If Obj.Selected Or Obj.EditPoint Then
          Obj.MakeTransformation
      End If
    Next Obj
End Sub

Public Property Let vbdObject_Name(ByVal RHS As String)
        m_Name = RHS
End Property

Public Property Get vbdObject_Name() As String
       vbdObject_Name = m_Name
End Property

Private Sub vbdObject_NewPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
            
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

Private Property Set vbdObject_Picture(ByVal RHS As stdole.Picture)
      Set m_Picture = RHS
End Property

Private Property Get vbdObject_Picture() As stdole.Picture
       Set vbdObject_Picture = m_Picture
End Property

Private Sub vbdObject_ReadPoint(ByRef NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
       ' m_Object.ReadPoint NumPoints, X(), Y(), TPoint()
End Sub

Private Sub vbdObject_ReadTrPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
          
End Sub

' Set the objects' Selected statuses.
Private Property Let vbdObject_Selected(ByVal RHS As Boolean)
    Dim Obj As vbdObject

    For Each Obj In SceneObjects
        Obj.Selected = RHS
    Next Obj
End Property
' Return the objects' Selected status.
Private Property Get vbdObject_Selected() As Boolean
    Dim Obj As vbdObject

    If SceneObjects.Count = 0 Then
        vbdObject_Selected = False
    Else
        Set Obj = SceneObjects(1)
        vbdObject_Selected = Obj.Selected
    End If
End Property

' Find the object at this position.
Public Function FindObjectAt(ByVal X As Single, ByVal Y As Single) As vbdObject
Dim Obj As vbdObject
Dim i As Integer

    Set FindObjectAt = Nothing

    ' Search for the object starting with
    ' the objects on top.
    For i = SceneObjects.Count To 1 Step -1
        Set Obj = SceneObjects(i)
        ''Debug.Print i, Obj.Serialization
        If PtInRegion(Obj.hRegion, X, Y) Then
            Set FindObjectAt = Obj
            Exit Function
        End If
    Next i
    
    For i = SceneObjects.Count To 1 Step -1
        Set Obj = SceneObjects(i)
        If Obj.IsAt(X, Y) Then
            Set FindObjectAt = Obj
            Exit For
        End If
    Next i
   ' FindObjectAt = Nothing
End Function

' Return True if the object is at this location.
Private Function vbdObject_IsAt(ByVal X As Single, ByVal Y As Single) As Boolean

End Function

' Initialize the object using a serialization string.
' The serialization does not include the
' ObjectType(...) part.
Private Property Let vbdObject_Serialization(ByVal RHS As String)
Dim Obj As vbdObject
Dim token_name As String
Dim token_value As String

    ' Remove non-printable characters from the serialization.
    RHS = RemoveNonPrintables(RHS)

    ' Start with no objects.
    Set SceneObjects = New Collection

    ' Read tokens until there are no more.
    Do While Len(RHS) > 0
        ' Read a token.
        GetNamedToken RHS, token_name, token_value
        Select Case token_name
            Case "Polygon", "RectAngle", "Picture"
                Set Obj = New vbdLine
            Case "Text", "TextFrame"
                Set Obj = New VbText
            Case "FreePolygon", "Polyline", "PolyDraw", "Curve"
                Set Obj = New vbdPolygon
            Case "Scribble"
                Set Obj = New vbdScribble
            Case "Scene"
                Set Obj = New vbdScene
            Case "Ellipse"
                Set Obj = New vbdEllipse
            'Case "Curve"
            '    Set Obj = New 'vbdCurve
            Case Else
                Set Obj = Nothing
        End Select

        ' Initialize the object.
        If Not Obj Is Nothing Then
            Obj.Serialization = token_value
            Obj.MakeTransformation
            SceneObjects.Add Obj
            Set Obj = Nothing
        End If
    Loop
End Property

' Return a serialization string for the object.
Private Property Get vbdObject_Serialization() As String
Dim txt As String
Dim Obj As vbdObject

    ' Don't bother with this object's
    ' drawing properties.

    ' Get the sub-objects' serializations.
    For Each Obj In SceneObjects
       ''Debug.Print Obj.TypeDraw
        txt = txt & vbCrLf & "  " & _
            Obj.Serialization
    Next Obj

    vbdObject_Serialization = _
        "Scene(" & txt & vbCrLf & _
        "  )"
End Property

Public Property Let vbdObject_Shade(ByVal RHS As Boolean)
        m_Shade = RHS
End Property

Public Property Get vbdObject_Shade() As Boolean
        vbdObject_Shade = m_Shade
End Property

Public Property Let vbdObject_Size(ByVal RHS As Integer)
        m_Size = RHS
End Property

Public Property Get vbdObject_Size() As Integer
       vbdObject_Size = m_Size
End Property

Private Sub vbdObject_StartBound(xmin As Single, ymin As Single)
     StartBoundObjects SceneObjects, xmin, ymin
End Sub

Public Property Let vbdObject_Strikethrough(ByVal RHS As Boolean)
          m_Strikethrough = RHS
End Property

Public Property Get vbdObject_Strikethrough() As Boolean
          vbdObject_Strikethrough = m_Strikethrough
End Property

Public Property Let vbdObject_TextDraw(ByVal new_value As String)
        m_TextDraw = new_value
End Property

Public Property Get vbdObject_TextDraw() As String
        vbdObject_TextDraw = m_TextDraw
End Property

' Set the object's TypeDraw.
Public Property Let vbdObject_TypeDraw(ByVal RHS As DrawType)
     m_TypeDraw = RHS
End Property
' Return the object's TypeDraw.
Public Property Get vbdObject_TypeDraw() As DrawType
     vbdObject_TypeDraw = m_TypeDraw
End Property

Public Property Let vbdObject_TypeFill(ByVal new_value As DrawTypeFill)
        m_TypeFill = new_value
End Property

Public Property Get vbdObject_TypeFill() As DrawTypeFill
       vbdObject_TypeFill = m_TypeFill
End Property


Public Property Let vbdObject_Underline(ByVal RHS As Boolean)
     m_Underline = RHS
End Property

Public Property Get vbdObject_Underline() As Boolean
         vbdObject_Underline = m_Underline
End Property

Public Property Let vbdObject_Weight(ByVal RHS As Integer)
        m_Weight = RHS
End Property

Public Property Get vbdObject_Weight() As Integer
       vbdObject_Weight = m_Weight
End Property


