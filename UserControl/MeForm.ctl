VERSION 5.00
Begin VB.UserControl MeForm 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   ControlContainer=   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   5145
   ToolboxBitmap   =   "MeForm.ctx":0000
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   3555
      Picture         =   "MeForm.ctx":0312
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   1905
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   3540
      Picture         =   "MeForm.ctx":0598
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   1635
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox ActiveTitle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   15
      ScaleHeight     =   225
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   0
      Width           =   5145
      Begin VB.PictureBox ButtonClose 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4905
         Picture         =   "MeForm.ctx":081E
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   2
         Top             =   15
         Width           =   195
         Begin VB.Line LineBorder 
            BorderColor     =   &H80000015&
            Index           =   0
            X1              =   0
            X2              =   45
            Y1              =   18
            Y2              =   18
         End
         Begin VB.Line LineBorder 
            BorderColor     =   &H80000015&
            Index           =   1
            X1              =   44
            X2              =   44
            Y1              =   0
            Y2              =   18
         End
         Begin VB.Line LineBorder 
            BorderColor     =   &H80000014&
            Index           =   2
            X1              =   0
            X2              =   44
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line LineBorder 
            BorderColor     =   &H80000014&
            Index           =   3
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   18
         End
         Begin VB.Line LineBorder 
            BorderColor     =   &H80000010&
            Index           =   4
            X1              =   43
            X2              =   43
            Y1              =   1
            Y2              =   19
         End
         Begin VB.Line LineBorder 
            BorderColor     =   &H80000010&
            Index           =   5
            X1              =   1
            X2              =   44
            Y1              =   17
            Y2              =   17
         End
      End
      Begin VB.Image ImageArrow 
         Height          =   165
         Left            =   4665
         Top             =   15
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   45
         Stretch         =   -1  'True
         Top             =   30
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MeForm"
         ForeColor       =   &H80000009&
         Height          =   165
         Left            =   300
         TabIndex        =   1
         Top             =   30
         Width           =   4215
      End
   End
   Begin VB.Image ImageArr 
      Height          =   195
      Index           =   1
      Left            =   4020
      Picture         =   "MeForm.ctx":0AA4
      Top             =   2010
      Width           =   195
   End
   Begin VB.Image ImageArr 
      Height          =   195
      Index           =   0
      Left            =   4095
      Picture         =   "MeForm.ctx":0D2A
      Top             =   1695
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   7
      X1              =   1395
      X2              =   915
      Y1              =   3210
      Y2              =   1965
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   6
      X1              =   1200
      X2              =   720
      Y1              =   3150
      Y2              =   1905
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   5
      X1              =   855
      X2              =   375
      Y1              =   1845
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   4
      X1              =   1050
      X2              =   570
      Y1              =   1905
      Y2              =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      Index           =   3
      X1              =   1065
      X2              =   585
      Y1              =   1845
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      Index           =   2
      X1              =   735
      X2              =   255
      Y1              =   1860
      Y2              =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   645
      X2              =   165
      Y1              =   1935
      Y2              =   690
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   1110
      X2              =   630
      Y1              =   1695
      Y2              =   450
   End
End
Attribute VB_Name = "MeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(c)2007 Diomidisk
'Event Declarations:
Event Hide() 'MappingInfo=ButtonClose,ButtonClose,-1,Click
Attribute Hide.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

'----<ãéá ãñáöéêÜ ìðáñáò
Const DC_ACTIVE = &H1   'ok
'Const DC_NOTACTIVE = &H2
Const DC_ICON = &H4  'ok
Const DC_TEXT = &H8  'ok
Const BDR_SUNKENOUTER = &H2
Const BDR_RAISEDINNER = &H4
Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Const BF_BOTTOM = &H8
Const BF_LEFT = &H1
Const BF_RIGHT = &H4
Const BF_TOP = &H2
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Const DC_GRADIENT = &H20    'ok      'Only Win98/2000 !!

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long, pcRect As RECT, ByVal un As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

'Default Property Values:
Const m_def_AutoRedraw = 0
Const m_def_MinButton = False
Const m_def_CloseButton = True
Const m_def_ObjHWND = -1

'Property Variables:
Dim m_AutoRedraw As Boolean
Dim m_MinButton As Boolean
Dim m_CloseButton As Boolean
Dim m_ObjHWND As Long

Private Sub ActiveTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     'RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub ActiveTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub ActiveTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      'RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub ButtonClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    LineBorder(0).BorderColor = vb3DHighlight
    LineBorder(1).BorderColor = vb3DHighlight
    LineBorder(2).BorderColor = vb3DDKShadow
    LineBorder(3).BorderColor = vb3DDKShadow
    LineBorder(4).BorderColor = vbButtonShadow
    LineBorder(5).BorderColor = vbButtonShadow
    ButtonClose.Picture = Picture1(1).Image
    
End Sub

Private Sub ButtonClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    LineBorder(0).BorderColor = vb3DDKShadow
    LineBorder(1).BorderColor = vb3DDKShadow
    LineBorder(2).BorderColor = vb3DHighlight
    LineBorder(3).BorderColor = vb3DHighlight
    LineBorder(4).BorderColor = vbButtonShadow
    LineBorder(5).BorderColor = vbButtonShadow
    ButtonClose.Picture = Picture1(0).Image
    
    If X > 0 And X < ButtonClose.Width Then
    If Y > 0 And Y < ButtonClose.Height Then
       RaiseEvent Hide
    End If
    End If
    
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub Image2_Click()

End Sub



Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     'RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'RaiseEvent MouseMove(Button, Shift, X, Y)
   UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     'RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
'
'Private Sub ButtonMax_Click()
'     RaiseEvent WinState(ButtonMax.WState)
'End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
    If Button = 1 Then
        'Release capture
        Call ReleaseCapture
        Dim TheHwnd As Long
        If m_ObjHWND <> -1 Then
            TheHwnd = m_ObjHWND
        Else
            TheHwnd = UserControl.hWnd ' .ContainerHwnd
        End If
        lngReturnValue = SendMessage(TheHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub UserControl_Resize()
      
      ActiveTitle.Move 30, 30, UserControl.Width - 60, 225
      
      If ButtonClose.Visible = True Then
         Label1.Width = ActiveTitle.Width - ButtonClose.Width
      Else
         Label1.Width = ActiveTitle.Width
 
      End If
      
      Label1.Move 30, (ActiveTitle.Height - Label1.Height) / 2
      ButtonClose.Top = (ActiveTitle.Height - ButtonClose.Height) / 2
      ButtonClose.Left = ActiveTitle.Width - ButtonClose.Width - 2 - 10
      
      DrawButton
      
      PaintControl
      
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

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
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    If New_BackStyle > 1 Then Exit Property
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Private Sub UserControl_Show()
    PaintControl
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000009)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Label1.Caption = PropBag.ReadProperty("Caption", "MeForm")
    m_CloseButton = PropBag.ReadProperty("CloseButton", m_def_CloseButton)
    Label1.Alignment = PropBag.ReadProperty("Alignment", Label1.Alignment)
    m_MinButton = PropBag.ReadProperty("MinButton", m_def_MinButton)
    m_MaxButton = PropBag.ReadProperty("MaxButton", m_def_MaxButton)
    m_HelpButton = PropBag.ReadProperty("HelpButton", m_def_HelpButton)
    m_ObjHWND = PropBag.ReadProperty("ObjHWND", m_def_ObjHWND)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
    m_AutoRedraw = PropBag.ReadProperty("AutoRedraw", m_def_AutoRedraw)
    Set Image1.Picture = PropBag.ReadProperty("Icon", Nothing)
    
    ButtonClose.Visible = m_CloseButton
    MinButton = m_MinButton
    DrawButton
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000009)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "MeForm")
    Call PropBag.WriteProperty("CloseButton", m_CloseButton, m_def_CloseButton)
    
    Call PropBag.WriteProperty("Alignment", Label1.Alignment, Label1.Alignment)
    Call PropBag.WriteProperty("MinButton", m_MinButton, m_def_MinButton)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, True)
    Call PropBag.WriteProperty("AutoRedraw", m_AutoRedraw, m_def_AutoRedraw)
    Call PropBag.WriteProperty("ObjHWND", m_ObjHWND, m_def_ObjHWND)

End Sub

Private Sub PaintControl()
     
     Line1(6).x1 = 0
     Line1(6).y1 = 0
     Line1(6).x2 = UserControl.Width
     Line1(6).y2 = 0
     
     Line1(7).x1 = 0
     Line1(7).y1 = 0
     Line1(7).x2 = 0
     Line1(7).y2 = UserControl.Height
     
     Line1(0).x1 = 20
     Line1(0).y1 = 20
     Line1(0).x2 = UserControl.Width
     Line1(0).y2 = 20
     
     Line1(1).x1 = 20
     Line1(1).y1 = 20
     Line1(1).x2 = 20
     Line1(1).y2 = UserControl.Height
     
     Line1(2).x1 = UserControl.Width - 20
     Line1(2).y1 = 0
     Line1(2).x2 = UserControl.Width - 20
     Line1(2).y2 = UserControl.Height - 20
     
     Line1(3).x1 = 0
     Line1(3).y1 = UserControl.Height - 20
     Line1(3).x2 = UserControl.Width - 20
     Line1(3).y2 = UserControl.Height - 20
     
     Line1(4).x1 = UserControl.Width - 30
     Line1(4).y1 = 30
     Line1(4).x2 = UserControl.Width - 30
     Line1(4).y2 = UserControl.Height - 30
     
     Line1(5).x1 = 30
     Line1(5).y1 = UserControl.Height - 30
     Line1(5).x2 = UserControl.Width - 20
     Line1(5).y2 = UserControl.Height - 30
    
    Dim R As RECT
    ActiveTitle.Cls
    UserControl.ScaleMode = vbPixels
    SetRect R, 0, 0, UserControl.ScaleWidth - 4, ActiveTitle.Height - 1
   
    DrawCaption ActiveTitle.hWnd, ActiveTitle.hDC, R, DC_ACTIVE Or DC_ICON Or DC_TEXT Or DC_GRADIENT
    DrawEdge ActiveTitle.hDC, R, EDGE_ETCHED, BF_RECT
    SetRect R, 0, 0, UserControl.ScaleWidth, ActiveTitle.Height
    DrawCaption UserControl.hWnd, ActiveTitle.hDC, R, DC_ACTIVE Or DC_ICON Or DC_TEXT Or DC_GRADIENT
    UserControl.ScaleMode = vbTwips
    
End Sub
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get CloseButton() As Boolean
    CloseButton = m_CloseButton
End Property
'
Public Property Let CloseButton(ByVal New_CloseButton As Boolean)
    
    m_CloseButton = New_CloseButton
    DrawButton
    ButtonClose.Visible = m_CloseButton
    'UserControl_Resize
    PropertyChanged "CloseButton"
    
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    m_CloseButton = m_def_CloseButton
    ButtonClose.Visible = m_CloseButton
    m_MinButton = m_def_MinButton
    m_MaxButton = m_def_MaxButton
    
    m_WState = m_def_WState
    m_AutoRedraw = m_def_AutoRedraw
    m_HelpButton = m_def_HelpButton

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = Label1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    Label1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Private Sub DrawButton()
     Dim lf As Long
     If Image1.Picture <> 0 Then
        lf = 300
     Else
        lf = 30
     End If
     ImageArrow.Picture = ImageArr(0).Picture
     If m_CloseButton Then
         ButtonClose.Move ActiveTitle.Width - ButtonClose.Width - 2 - 10
         ImageArrow.Move ActiveTitle.Width - (ButtonClose.Width) - 50 - ImageArrow.Width, ButtonClose.Top
         If m_MaxButton Then
            ButtonMax.Move ActiveTitle.Width - ButtonClose.Width - ButtonMax.Width - 4 - 10, (ActiveTitle.Height - ButtonMax.Height) / 2
            ButtonMin.Move ActiveTitle.Width - ButtonClose.Width - ButtonMax.Width - ButtonMin.Width - 6 - 10, (ActiveTitle.Height - ButtonMin.Height) / 2
            Buttonhelp.Move ActiveTitle.Width - ButtonClose.Width - ButtonMax.Width - ButtonMin.Width - Buttonhelp.Width - 8 - 10, (ActiveTitle.Height - ButtonMin.Height) / 2
            If m_MinButton Then
               Label1.Move lf, (ActiveTitle.Height - Label1.Height) / 2, (ActiveTitle.Width - ButtonClose.Width - ButtonMax.Width - ButtonMin.Width) - 60
            Else
               Label1.Move lf, (ActiveTitle.Height - Label1.Height) / 2, (ActiveTitle.Width - ButtonClose.Width - ButtonMax.Width) - 60
            End If
            
         Else
            If m_MinButton Then
               Label1.Move lf, (ActiveTitle.Height - Label1.Height) / 2, (ActiveTitle.Width - ButtonClose.Width - ButtonMin.Width) - 60
            Else
               Label1.Move lf, (ActiveTitle.Height - Label1.Height) / 2, (ActiveTitle.Width - ButtonClose.Width) - 60
            End If
            
         End If
         
         
      End If
      
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,
Public Property Get WState() As Integer
    WState = m_WState
End Property

Public Property Let WState(ByVal New_WState As Integer)
    m_WState = New_WState
    PropertyChanged "WState"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = m_AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    m_AutoRedraw = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Picture
Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Icon = Image1.Picture
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set Image1.Picture = New_Icon
    PropertyChanged "Icon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,2,-1
Public Property Get ObjHWND() As Long
    ObjHWND = m_ObjHWND
End Property

Public Property Let ObjHWND(ByVal New_ObjHWND As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    m_ObjHWND = New_ObjHWND
    PropertyChanged "ObjHWND"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

