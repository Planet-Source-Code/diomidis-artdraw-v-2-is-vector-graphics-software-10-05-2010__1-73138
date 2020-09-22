Attribute VB_Name = "ModVbDraw"

' The new object we are building.
Public m_NewObject As vbdObject
Public m_EditObject As vbdObject

' The selected object.
Public m_SelectedObjects As Collection

' Undo variables.
Public Const MAX_UNDO = 500
Public Const GAP = 6
Public gZoomFactor As Single

Public m_Snapshots As Collection
Public m_CurrentSnapshot As Integer

' The scene that holds all objects.
Public m_TheScene As vbdObject

' Current world window bounds.
Public Wxmin As Long
Public Wxmax As Long
Public Wymin As Long
Public Wymax As Long

'Public gZoomLock As Boolean
Public gScaleMode As Integer
Public gPrintetOrientation As Integer
Public m_FormSymbolView As Boolean
Public m_FormMagnify As Boolean

' Indicates the data has changed since load/save.
Public m_DataModified As Boolean
Public m_ViewSimple As Boolean
'
Public OldMode As Long
Public OldOrg As PointAPI
Public WndArea As RECT
'

Public Enum DrawType
    [dPolyline] = 0
    [dScribble] = 1
    [dFreePolygon] = 2
    [dPolygon] = 3
    [dRectAngle] = 4
    [dEllipse] = 5
    [dText] = 6
    [dTextArt] = 7
    [dTextPath] = 8
    [dPolydraw] = 9
    [dPicture] = 10
    [dTextFrame] = 11
    [dCurve] = 12
    [dCalligraphic] = 13
End Enum
Public Enum DrawTypeFill
    [dSimple] = 0
    [dBitmap] = 1
    [dGradient] = 2
End Enum

Public Const PI = 3.14159265358979


Public Sub Main()
     Load frmAbout1
     frmAbout1.Show
        frmVbDraw.Show
     Unload frmAbout1
End Sub
' Add this object to the collection.
Public Sub AddObject(ByVal Obj As vbdObject) ', _
'                     ByVal Fore_Color As Integer, _
'                     ByVal Fill_Color As Long, _
'                     ByVal Draw_Style As Integer, _
'                     ByVal Fill_Style As Integer, _
'                     ByVal Draw_Width As Integer, _
'                     ByVal Type_Draw As Integer, _
'                     ByVal TypeFill As Integer, _
'                     ByVal tShade As Boolean, _
'                     ByVal Text_Draw As String)
                     
                     
Dim the_scene As vbdScene

    ' Give the object its drawing properties.
  '  obj.ForeColor = QBColor(m_ForeColor)
  '  obj.FillColor = QBColor(m_BackColor)
  '  obj.DrawStyle = icbDrawStyle.SelectedItem.Index - 1
  '  obj.FillStyle = icbFillStyle.SelectedItem.Index - 1
  '  Obj.DrawWidth = icbDrawWidth.SelectedItem.Index
     
    ' Save the new object.
    Set the_scene = m_TheScene
    the_scene.SceneObjects.Add Obj
    'the_scene.vbdObject_TypeDraw = Obj.TypeDraw
    'the_scene.vbdObject_TypeFill = Obj.TypeFill
    'the_scene.vbdObject_Shade = Obj.Shade
    'the_scene.vbdObject_TextDraw = Obj.TextDraw
    Set m_NewObject = Nothing
    
    ' Select the new object only.
    DeselectAllVbdObjects
   ' SelectVbdObject Obj

    ' See if any objects are selected.
    frmVbDraw.DrawControl1_EnableMenusForSelection

    ' Select the arrow tool.
    frmVbDraw.SelectArrowTool

    ' Save the current snapshot.
     SaveSnapshot
    
    'Redraw.
    frmVbDraw.DrawControl1.Redraw
End Sub
  
'Public Sub CancelEditPoint()
'       frmVbDraw.DrawControl1.EditPoint = False
'       frmVbDraw.drawToolbar.CheckButton 1, True
'End Sub

' Deselect all objects.
Public Sub DeselectAllVbdObjects()
Dim Obj As vbdObject

    ' Deselect all selected objects.
    For Each Obj In m_SelectedObjects
       If Obj.Selected Then
          Obj.Selected = False
       End If
       If Obj.EditPoint Then
          Obj.EditPoint = False
       End If
       
    Next Obj

    ' Empty the m_SelectedObjects collection.
    Set m_SelectedObjects = New Collection
End Sub
' Select all objects.
Public Sub SelectAllVbdObjects()
Dim Obj As vbdObject

    ' Deselect all selected objects.
    For Each Obj In m_SelectedObjects
        Obj.Selected = True
    Next Obj

    ' Empty the m_SelectedObjects collection.
    Set m_SelectedObjects = New Collection
End Sub


' Select this object.
Public Sub SelectVbdObject(ByVal target As vbdObject)
    ' See if it is aleady selected.
    If target.Selected Then Exit Sub
     
    ' Add the object to the
    ' m_SelectedObjects collection.
    m_SelectedObjects.Add target
   
    ' Mark the object as selected.
    target.Selected = True
End Sub

' Select this object.
Public Sub SelectVbPoint(ByVal target As vbdObject)
    ' See if it is aleady selected.
    If target.EditPoint Then Exit Sub
     
    ' Add the object to the
    ' m_SelectedObjects collection.
    m_SelectedObjects.Add target
   
    ' Mark the object as selected.
    target.EditPoint = True
End Sub

Public Sub CancelObject()
    Set m_NewObject = Nothing

    ' Select the arrow tool.
    frmVbDraw.SelectArrowTool
End Sub

' Deselect this object.
Public Sub DeselectVbdObject(ByVal target As vbdObject)
    Dim Obj As vbdObject
    Dim i As Integer

    ' Remove the object from the
    ' m_SelectedObjects collection.
    i = 1
    For Each Obj In m_SelectedObjects
        If Obj Is target Then
            m_SelectedObjects.Remove i
            Exit For
        End If
        i = i + 1
    Next Obj

    ' Mark the object as not selected.
    target.Selected = False
End Sub

Public Sub DeletevbdObject()
    Dim the_scene As vbdScene
    If m_SelectedObjects.Count > 0 Then
       ' Delete the selected objects.
        Set the_scene = m_TheScene
         the_scene.RemoveObjects m_SelectedObjects
       End If
End Sub
' Find the object at this position.
Public Function FindObjectAt(ByVal X As Single, ByVal Y As Single) As vbdObject
    Dim the_scene As vbdScene, Fobj As vbdObject
    Set the_scene = m_TheScene
    Set Fobj = the_scene.FindObjectAt(X, Y)
    If Not Fobj Is Nothing Then
       Set FindObjectAt = Fobj
    End If
End Function

'' Delete the object.
'Public Sub DeleteObj()
'    Dim the_scene As vbdScene
'       If m_SelectedObjects.Count > 0 Then
'            ' Delete the selected objects.
'            Set the_scene = m_TheScene
'            the_scene.RemoveObjects m_SelectedObjects
'            ' The data has changed.
'            ' Save the current snapshot.
'            SaveSnapshot
'            PicCanvas.Refresh
'       End If
'End Sub

' Save a snapshot for undo.
Public Sub SaveSnapshot()
    
    If m_Snapshots Is Nothing Then Exit Sub
    ' Remove any previously undone snapshots.
    Do While m_Snapshots.Count > m_CurrentSnapshot
        m_Snapshots.Remove m_Snapshots.Count
    Loop

    ' Save the current snapshot.
    m_Snapshots.Add m_TheScene.Serialization
    ''Debug.Print TmpSer
    If m_Snapshots.Count > MAX_UNDO Then
        For i = 1 To 50
        m_Snapshots.Remove 1
        Next
    End If
    m_CurrentSnapshot = m_Snapshots.Count

End Sub

Public Sub ChangeFillColor(IdColor As Integer, mColor As Long)
    Dim Obj As vbdObject

    ' Change all selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True Or Obj.EditPoint = True Then
           If IdColor = 1 Then
           Obj.FillColor = mColor
           Else
           Obj.FillColor2 = mColor
           End If
        End If
    Next Obj

End Sub

Public Sub ChangeForeColor(mColor As Long)
    Dim Obj As vbdObject
    
    ' Change all selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True Or Obj.EditPoint = True Then
           Obj.ForeColor = mColor
        End If
    Next Obj
End Sub

Public Sub ChangeFillstyle(mFillStyle As Integer)
    Dim Obj As vbdObject
    
    ' Change all selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True Or Obj.EditPoint = True Then
           Obj.FillStyle = mFillStyle
        End If
    Next Obj
End Sub

Public Sub ChangeBlend(mBlend As Integer)
    Dim Obj As vbdObject
    
    ' Change all selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True Or Obj.EditPoint = True Then
           Obj.Blend = mBlend
        End If
    Next Obj
End Sub

Public Sub ChangeDrawWidth(mDrawWidth As Integer)
    Dim Obj As vbdObject
    
    ' Change all selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True Or Obj.EditPoint = True Then
           Obj.DrawWidth = mDrawWidth
        End If
    Next Obj
End Sub

Public Sub ChangeDrawstyle(mDrawStyle As Integer)
    Dim Obj As vbdObject
    
    ' Change all selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True Or Obj.EditPoint = True Then
           Obj.DrawStyle = mDrawStyle
        End If
    Next Obj
End Sub

Public Sub ChangePattern(nPattern As String)
    Dim Obj As vbdObject

    ' Change all selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True Or Obj.EditPoint = True Then
           Obj.Pattern = nPattern
        End If
    Next Obj

End Sub

Public Sub ChangeGradient(nId As Integer)
    Dim Obj As vbdObject

    ' Change all selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True Or Obj.EditPoint = True Then
           Obj.Gradient = nId
        End If
    Next Obj

End Sub

Public Function FindObject(txt As String) As Boolean
       If InStr(txt, "Polygon") > 0 Or _
          InStr(txt, "Polyline") > 0 Or _
          InStr(txt, "Scribble") > 0 Or _
          InStr(txt, "FreePolygon") > 0 Or _
          InStr(txt, "RectAngle") > 0 Or _
          InStr(txt, "Picture") > 0 Or _
          InStr(txt, "PolyDraw") > 0 Or _
          InStr(txt, "Ellipse") > 0 Or _
          InStr(txt, "TextFrame") > 0 Or _
          InStr(txt, "TextPath") > 0 Or _
          InStr(txt, "Text") > 0 Then
          FindObject = True
       Else
          FindObject = False
       End If
End Function

Public Sub OpenColorDialog(pic As PictureBox)
  '  Dim cf As ColorDialog
    Dim nColor As Long
    Dim C As SelectedColor
    nColor = pic.BackColor
     ColorDialog.rgbResult = nColor
    C = ShowColor(pic.hWnd, False)
    If C.bCanceled = False Then
      pic.BackColor = C.oSelectedColor
    End If
End Sub


Private Function RectWidth(ByRef inRect As RECT) As Long
    RectWidth = inRect.Right - inRect.Left   ' Width of a rectangle
End Function

Private Function RectHeight(ByRef inRect As RECT) As Long
    RectHeight = inRect.Bottom - inRect.Top ' Height of a rectangle
End Function

Private Function RectExtent(ByRef inRect As RECT) As SIZEL
    RectExtent.cx = RectWidth(inRect) ' Extent of a rectangle
    RectExtent.cy = RectHeight(inRect)
End Function

Private Function RectOrigin(ByRef inRect As RECT) As PointAPI
    RectOrigin.X = inRect.Left ' Origin of a rectangle
    RectOrigin.Y = inRect.Top
End Function

Private Function CreateScaleMode(ByRef inSource As RECT, ByRef inDest As RECT, _
                                Optional ByVal inIsotropic As Boolean = True) As typScaleMode
    With CreateScaleMode ' Creates mapping between two rectangles
        .smViewportExtent = RectExtent(inSource)
        .smViewportOrigin = RectOrigin(inSource)
        .smWindowExtent = RectExtent(inDest)
        .smWindowOrigin = RectOrigin(inDest)
        .smMappingMode = IIf(inIsotropic, MM_ISOTROPIC, MM_ANISOTROPIC)
    End With
End Function

Private Function SetScaleMode(ByVal inDC As Long, ByRef inScaleMode As typScaleMode) As typScaleMode
                            
    With inScaleMode ' Applies a scale mode to a DC and returns the old one
        SetScaleMode.smMappingMode = SetMapMode(inDC, .smMappingMode)
        Call SetWindowOrgEx(inDC, .smWindowOrigin.X, .smWindowOrigin.Y, SetScaleMode.smWindowOrigin)
        Call SetWindowExtEx(inDC, .smWindowExtent.cx, .smWindowExtent.cy, SetScaleMode.smWindowExtent)
        Call SetViewportOrgEx(inDC, .smViewportOrigin.X, .smViewportOrigin.Y, SetScaleMode.smViewportOrigin)
        Call SetViewportExtEx(inDC, .smViewportExtent.cx, .smViewportExtent.cy, SetScaleMode.smViewportExtent)
    End With
End Function

Public Function SetWordScale(canvas As PictureBox) As typScaleMode
    Dim OldScale As typScaleMode, NewScale As typScaleMode
    Dim ClientArea As RECT, ScaleArea As RECT
    
    ' Get window size and trim 20 pixel border as source area
    Call GetClientRect(canvas.hWnd, ClientArea)

    ' Set mapping area with central origin
    Call SetRect(ScaleArea, Wxmin, Wymin, Wxmax, Wymax)
   
   ' Create scale mode and set up DC
    NewScale = CreateScaleMode(ClientArea, ScaleArea)
  '  Call DPtoLP(inDC, ScaleArea, 2)
    OldScale = SetScaleMode(canvas.hDC, NewScale)
     
    SetWordScale = OldScale
End Function

Public Sub ResetWordScale(canvas As PictureBox, OldScale As typScaleMode)
       ' Restore old mapping mode
       Call SetScaleMode(canvas.hDC, OldScale)
End Sub

Public Sub SetDC(ByVal inDC As Long, ByVal inWnd As Long, xmin As Long, ymin As Long)
    ' Find the size of the window, and fill it with the background colour
  '  Call GetClientRect(inWnd, WndArea)
   ' Call FillRect(inDC, WndArea, GetSysColorBrush(COLOR_BTNFACE))
 '   Call OffsetRect(WndArea, -WndArea.Left, -WndArea.Top)

    ' Set the mapping mode and centre origin
  '  OldMode = SetMapMode(inDC, MM_TWIPS)
  '  Call DPtoLP(inDC, WndArea, 2)
    Call SetWindowOrgEx(inDC, xmin, ymin, OldOrg)
End Sub

Public Sub ClearDC(ByVal inDC As Long)
    ' Re-set the window origin and mapping mode to keep VB happy
   ' Call SetWindowOrgEx(inDC, OldOrg.X, OldOrg.Y, ByVal 0&)
    'Call SetMapMode(inDC, OldMode)
End Sub
