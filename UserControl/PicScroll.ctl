VERSION 5.00
Begin VB.UserControl PicScroll 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   ScaleHeight     =   6780
   ScaleWidth      =   9360
   Begin VB.PictureBox Grabber 
      BackColor       =   &H80000000&
      Height          =   255
      Left            =   8895
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   6450
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picShow 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   555
      ScaleHeight     =   5415
      ScaleWidth      =   7020
      TabIndex        =   2
      Top             =   540
      Width           =   7020
      Begin VB.PictureBox picNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9495
         Left            =   1740
         ScaleHeight     =   9495
         ScaleWidth      =   12615
         TabIndex        =   6
         Top             =   3450
         Visible         =   0   'False
         Width           =   12615
      End
      Begin VB.PictureBox picHold 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1815
         Left            =   135
         ScaleHeight     =   1815
         ScaleWidth      =   2175
         TabIndex        =   3
         Top             =   90
         Width           =   2175
         Begin VB.PictureBox picDoc 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   1605
            Left            =   105
            ScaleHeight     =   1605
            ScaleWidth      =   1965
            TabIndex        =   4
            Top             =   105
            Visible         =   0   'False
            Width           =   1965
         End
      End
   End
   Begin VB.HScrollBar hsPreview 
      Height          =   255
      Left            =   6735
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.VScrollBar vsPreview 
      Height          =   1215
      Left            =   8895
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5220
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "PicScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_Zoom = 100
'Property Variables:
Dim m_Zoom As Integer

Private bScrollCode As Boolean
Private sZoom As Single
Private lPage As Integer
Private lPageMax As Integer
Private bDisplayPage As Boolean

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,100
Public Property Get Zoom() As Integer
    Zoom = m_Zoom
End Property

Public Property Let Zoom(ByVal New_Zoom As Integer)
    m_Zoom = New_Zoom
    Zoom_Check
    PropertyChanged "Zoom"
End Property

Private Sub hsPreview_Change()
If Not bScrollCode Then
  picHold.Left = -hsPreview.Value * 14.4
End If

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Zoom = m_def_Zoom
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set picNormal.Picture = Nothing  'PropBag.ReadProperty("Picture", Nothing)
    m_Zoom = PropBag.ReadProperty("Zoom", m_def_Zoom)
    picDoc.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 100)
    picDoc.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 100)
    
    vsPreview.Move UserControl.ScaleWidth - vsPreview.Width, 0, vsPreview.Width, UserControl.ScaleHeight - hsPreview.Height
    hsPreview.Move 0, UserControl.ScaleHeight - hsPreview.Height, UserControl.ScaleWidth - vsPreview.Width
    picDoc.Move -picDoc.Width, -picDoc.Height
    sZoom = 100
    m_Zoom = 100

End Sub

Private Sub UserControl_Resize()
    vsPreview.Left = UserControl.ScaleWidth - vsPreview.Width
    vsPreview.Height = UserControl.ScaleHeight
    hsPreview.Top = UserControl.ScaleHeight - hsPreview.Height
    hsPreview.Width = UserControl.ScaleWidth
    Zoom_Check
End Sub

Private Sub UserControl_Show()
'Me.Refresh
bDisplayPage = True
Preview_Display 1
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Picture", picNormal.Picture, Nothing)
    Call PropBag.WriteProperty("Zoom", m_Zoom, m_def_Zoom)
    Call PropBag.WriteProperty("ScaleHeight", picDoc.ScaleHeight, 100)
    Call PropBag.WriteProperty("ScaleWidth", picDoc.ScaleWidth, 100)
End Sub

Private Sub Zoom_Check()

Dim sSizeX As Single
Dim sSizeY As Single
Dim sRatio As Single
Dim spImage As StdPicture
Dim sWidth As Single
Dim sHeight As Single
Dim bScroll As Byte
Dim bOldScroll As Byte
Screen.MousePointer = vbHourglass

    sWidth = UserControl.ScaleWidth
    sHeight = UserControl.ScaleHeight
   sZoom = m_Zoom
    Do
        bOldScroll = bScroll
        If sZoom = 0 Then
            sRatio = (sHeight - 480) / picNormal.Height
        ElseIf sZoom = -1 Then
            sRatio = (sWidth - 480) / picNormal.Width
        Else
            sRatio = sZoom / 100
        End If
        sSizeX = picNormal.Width * sRatio
        sSizeY = picNormal.Height * sRatio
        If sSizeX > sWidth And (bScroll And 1) <> 1 Then
            sHeight = sHeight - hsPreview.Height
            bScroll = bScroll + 1
        End If
        If sSizeY > sHeight And (bScroll And 2) <> 2 Then
            sWidth = sWidth - vsPreview.Width
            bScroll = bScroll + 2
        End If
    Loop While bOldScroll <> bScroll

    vsPreview.Height = sHeight
    hsPreview.Width = sWidth

    picShow.Move 0, 0, sWidth, sHeight
    picDoc.Move 0, 0, sSizeX + 480, sSizeY + 480 '240, 240, sSizeX, sSizeY
    picDoc.Cls
    If picNormal.Picture = 0 Then Exit Sub
    picDoc.PaintPicture picNormal.Image, 0, 0, sSizeX, sSizeY
    bScrollCode = True
    picHold.Move 0, 0, sSizeX + 480, sSizeY + 480
    
    If (bScroll And 2) = 2 Then
        vsPreview.Visible = True
        vsPreview.Max = (picHold.ScaleHeight - picShow.ScaleHeight) / 14.4 + 1
        vsPreview.Min = 0
        vsPreview.SmallChange = 14
        vsPreview.LargeChange = picShow.ScaleHeight / 14.4
        vsPreview.Value = vsPreview.Min
    Else
        vsPreview.Visible = False
    End If

    If (bScroll And 1) = 1 Then
        hsPreview.Visible = True
        hsPreview.Max = (picHold.ScaleWidth - picShow.ScaleWidth) / 14.4 + 1
        hsPreview.Min = 0
        hsPreview.SmallChange = 14
        hsPreview.LargeChange = picShow.ScaleWidth / 14.4
        hsPreview.Value = hsPreview.Min
    Else
        hsPreview.Visible = False
    End If
    bScrollCode = False
    Screen.MousePointer = vbDefault
    If bDisplayPage Then
        picDoc.Visible = True
    End If
    
End Sub

Private Sub vsPreview_Change()
  If Not bScrollCode Then
        picHold.Top = -vsPreview.Value * 14.4
    End If
End Sub

Public Sub Preview_Display(ByVal iPage As Integer)

    Dim iMin As Integer
    Dim iMax As Integer
        Screen.MousePointer = vbHourglass
        picNormal.Cls
        Zoom_Check
        Screen.MousePointer = vbDefault
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picDoc,picDoc,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = picDoc.Image
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picDoc,picDoc,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    picDoc.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picDoc,picDoc,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = picDoc.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picDoc,picDoc,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = picDoc.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    picDoc.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picDoc,picDoc,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = picDoc.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    picDoc.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picNormal,picNormal,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = picNormal.Picture
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set picNormal.Picture = New_Picture
    'Set picNormal.Image = picNormal.Picture
    PropertyChanged "Picture"
    Zoom_Check
End Property

