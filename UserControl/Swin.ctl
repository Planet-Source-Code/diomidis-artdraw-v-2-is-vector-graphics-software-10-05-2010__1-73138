VERSION 5.00
Begin VB.UserControl ScrolledPicture 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   ControlContainer=   -1  'True
   ScaleHeight     =   2175
   ScaleWidth      =   2310
   ToolboxBitmap   =   "Swin.ctx":0000
   Begin VB.HScrollBar hbar 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.VScrollBar vbar 
      Height          =   1335
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picOuter 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      Begin VB.PictureBox picInner 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   855
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "ScrolledPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type PanState
   X As Long
   Y As Long
End Type
Dim PanSet As PanState

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Scroll(HValue As Integer, Vvalue As Integer)

' Reposition picInner.
Private Sub hbar_Change()
    picInner.Left = -hbar.Value
    EventStroll
End Sub

' Reposition picInner.
Private Sub hbar_Scroll()
    picInner.Left = -hbar.Value
    EventStroll
End Sub

' Reparent the contained controls into picInner and see how much room they need.
Private Sub ReparentControls()
Dim ctl As Control
Dim xmax As Single
Dim ymax As Single
Dim need_wid As Single
Dim need_hgt As Single

    ' Do nothing if no controls have been loaded.
    If ContainedControls.Count <> 0 Then

        For Each ctl In ContainedControls
            SetParent ctl.hWnd, picInner.hWnd

            xmax = ctl.Left + ctl.Width
            ymax = ctl.Top + ctl.Height
            If need_wid < xmax Then need_wid = xmax
            If need_hgt < ymax Then need_hgt = ymax
        Next ctl
        
          ' Make picInner big enough to hold the controls.
          picInner.Move 0, 0, need_wid, need_hgt
        
    ElseIf picInner.Picture <> 0 Then
        If picInner.Width > UserControl.Width Then
           picInner.Move 0, 0
        Else
          picInner.Move (UserControl.Width - picInner.Width) / 2, (UserControl.Height - picInner.Height) / 2   ', need_wid, need_hgt
        End If
    Else
       Exit Sub
    End If
    ' Hide the borders on picInner and picOuter.
    picOuter.BorderStyle = vbBSNone
    picInner.BorderStyle = vbBSNone
End Sub



Private Sub picInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Local Error Resume Next
   If Button = vbLeftButton And Shift = 0 Then
      PanSet.X = X
      PanSet.Y = Y
      MousePointer = vbSizePointer
   End If
End Sub

Private Sub picInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim nTop As Integer, nLeft As Integer

   On Local Error Resume Next

   If Button = vbLeftButton And Shift = 0 Then

      '/* new coordinates?
      With picInner
         nTop = -(.Top + (Y - PanSet.Y))
         nLeft = -(.Left + (X - PanSet.X))
      End With

      '/* Check limits
      With vbar
         If .Visible Then
            If nTop < .Min Then
               nTop = .Min
            ElseIf nTop > .Max Then
               nTop = .Max
            End If
         Else
            nTop = -picInner.Top
         End If
      End With

      With hbar
         If .Visible Then
            If nLeft < .Min Then
               nLeft = .Min
            ElseIf nLeft > .Max Then
               nLeft = .Max
            End If
         Else
            nLeft = -picInner.Left
         End If
      End With

      picInner.Move -nLeft, -nTop

   End If
End Sub

Private Sub picInner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Local Error Resume Next
   If Button = vbLeftButton And Shift = 0 Then
      If vbar.Visible Then vbar.Value = -(picInner.Top)
      If hbar.Visible Then hbar.Value = -(picInner.Left)
   End If
   MousePointer = vbDefault
End Sub

Private Sub UserControl_Resize()
    ' Hide the controls at design time.
    If Not Ambient.UserMode Then
        vbar.Visible = False
        hbar.Visible = False
        picInner.Visible = False
        Exit Sub
    End If

    ' Arrange the controls.
    ArrangeControls
End Sub

' Arrange the scroll bars.
Private Sub ArrangeControls()
Dim border_width As Single
Dim got_wid As Single
Dim got_hgt As Single
Dim need_wid As Single
Dim need_hgt As Single
Dim need_hbar As Boolean
Dim need_vbar As Boolean

    ' Reparent the controls.
    ReparentControls

    ' See how much room we have and need.
    border_width = picOuter.Width - picOuter.ScaleWidth
    got_wid = ScaleWidth - border_width
    got_hgt = ScaleHeight - border_width
    need_wid = picInner.Width
    need_hgt = picInner.Height

    ' See if we need the horizontal scroll bar.
    If need_wid > got_wid Then
        need_hbar = True
        got_hgt = got_hgt - hbar.Height
    End If

    ' See if we need the vertical scroll bar.
    If need_hgt > got_hgt Then
        need_vbar = True
        got_wid = got_wid - vbar.Width

        ' See if we now need the horizontal scroll bar.
        If (Not need_hbar) And need_wid > got_wid Then
            need_hbar = True
            got_hgt = got_hgt - hbar.Height
        End If
    End If

    ' Arrange the controls.
    picOuter.Move 0, 0, got_wid + border_width, got_hgt + border_width
    If need_hbar Then
        hbar.Move 0, got_hgt + border_width, got_wid + border_width
        hbar.Min = 0
        hbar.Max = picInner.ScaleWidth - got_wid
        hbar.SmallChange = got_wid / 5
        hbar.LargeChange = got_wid
        hbar.Visible = True
    Else
        hbar.Value = 0
        hbar.Visible = False
    End If
    If need_vbar Then
        vbar.Move got_wid + border_width, 0, vbar.Width, got_hgt + border_width
        vbar.Min = 0
        vbar.Max = picInner.ScaleHeight - got_hgt
        vbar.SmallChange = got_hgt / 5
        vbar.LargeChange = got_hgt
        vbar.Visible = True
    Else
        vbar.Value = 0
        vbar.Visible = False
    End If
End Sub

' Reposition picInner.
Private Sub vbar_Change()
    picInner.Top = -vbar.Value
    EventStroll
End Sub

' Reposition picInner.
Private Sub vbar_Scroll()
    picInner.Top = -vbar.Value
    EventStroll
End Sub

Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl_Resize
    UserControl.Refresh
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    hbar.Value = PropBag.ReadProperty("Hscroll", 0)
    vbar.Value = PropBag.ReadProperty("Vscroll", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Hscroll", hbar.Value, 0)
    Call PropBag.WriteProperty("Vscroll", vbar.Value, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

Public Property Get Hscroll() As Integer
Attribute Hscroll.VB_Description = "Returns/sets the value of an object."
    Hscroll = hbar.Value
End Property

Public Property Let Hscroll(ByVal New_Hscroll As Integer)
    hbar.Value() = New_Hscroll
    PropertyChanged "Hscroll"
End Property

Public Property Get Vscroll() As Integer
Attribute Vscroll.VB_Description = "Returns/sets the value of an object."
    Vscroll = vbar.Value
End Property

Public Property Let Vscroll(ByVal New_Vscroll As Integer)
    vbar.Value() = New_Vscroll
    PropertyChanged "Vscroll"
End Property

Private Sub EventStroll()
    RaiseEvent Scroll(hbar.Value, vbar.Value)
End Sub

Public Sub ReadScroll(HValue As Integer, Vvalue As Integer)
       HValue = hbar.Value
       Vvalue = vbar.Value
End Sub

Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = picInner.Picture
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    If New_Picture Is Nothing Then Exit Property
    Set picInner.Picture = New_Picture
    picInner.Refresh
''  picInner.Width = New_Picture.Width
''  picInner.Height = New_Picture.Height
    
    PropertyChanged "Picture"
End Property

