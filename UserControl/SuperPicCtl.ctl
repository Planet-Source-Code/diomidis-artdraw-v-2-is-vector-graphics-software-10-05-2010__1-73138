VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl SuperPicCtl 
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   Begin MSComctlLib.Toolbar tbrQuickFloat 
      Height          =   330
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlTools"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer timToolbar 
      Interval        =   100
      Left            =   1125
      Top             =   4560
   End
   Begin MSComctlLib.ImageList imlTools 
      Left            =   120
      Top             =   4485
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SuperPicCtl.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SuperPicCtl.ctx":01DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   4605
      Left            =   375
      ScaleHeight     =   303
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   415
      TabIndex        =   1
      Top             =   195
      Width           =   6285
      Begin VB.PictureBox picTemp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5925
         ScaleHeight     =   375
         ScaleWidth      =   345
         TabIndex        =   4
         Top             =   4245
         Width           =   345
      End
      Begin VB.HScrollBar hscPicture 
         Height          =   255
         LargeChange     =   20
         Left            =   -15
         SmallChange     =   5
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4290
         Visible         =   0   'False
         Width           =   5985
      End
      Begin VB.VScrollBar vscPicture 
         Height          =   4305
         LargeChange     =   20
         Left            =   5970
         SmallChange     =   5
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picView 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   0
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   129
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   6675
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   6
      Top             =   4785
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2235
      Picture         =   "SuperPicCtl.ctx":03B4
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1815
      Picture         =   "SuperPicCtl.ctx":0C7E
      Top             =   4815
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "SuperPicCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************************************************'
' Module    : SuperPicCtl
' Author    : Joseph M. Ferris <josephmferris@cox.net>
' Date      : 01.27.2004
' Depends   : Visual Basic Profession 6.0, SP5
' Purpose   : Provide a drop-in PictureBox replacement that adds Zoom support and native scrolling
' Notes     : 1.  Based on code submission to PlanetSourceCode from "amirnezhad".
'                 (http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=28511&lngWId=1)
'             2.  Added the StorePosition() and RecallPosition() methods to allow for restoring cursor positions on
'                 state changes (02.05.2004)
'             3.  Added the PanActive() property which toggles between standard and panning modes (02.05.2004)
'             4.  Added the Picture() property to assign and retrieve images to a StdPicture (02.06.2004)
'
'**********************************************************************************************************************'
Option Explicit
Private Const WM_USER                   As Long = &H400
Private Const TB_GETBUTTON              As Long = (WM_USER + 23)
Private Const SM_CXBORDER               As Long = 5                ' Width of non-sizable borders
Private Const SM_CYBORDER               As Long = 6                ' Height of non-sizable borders
Private Const SM_CXDLGFRAME             As Long = 7                ' Width of dialog box borders
Private Const SM_CYDLGFRAME             As Long = 8                ' Height of dialog box borders
Private Const SM_CYVTHUMB               As Long = 9                ' Height of scroll box on vertical scroll bar
Private Const SM_CXHTHUMB               As Long = 10               ' Width of scroll box on horizontal scroll bar

Private Type OriginalImage
    Height                              As Long
    Width                               As Long
End Type

Private Type POINTAPI
    X                                   As Long
    Y                                   As Long
End Type

Private Type TBBUTTON
   iBitmap                              As Long
   idCommand                            As Long
   fsState                              As Byte
   fsStyle                              As Byte
   bReserved1                           As Byte
   bReserved2                           As Byte
   dwData                               As Long
   iString                              As Long
End Type

Private Type ScrollPositions
    HorizontalScrollMax                 As Long
    HorizontalScrollPosition            As Long
    VerticalScrollMax                   As Long
    VerticalScrollPosition              As Long
End Type

Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private m_dblPercentage                 As Double
Private m_strFileName                   As String
Private m_udtOriginal                   As OriginalImage
Private m_bolAllowOut                   As Boolean
Private m_bolAllowIn                    As Boolean
Private m_bolUseQuickBar                As Boolean
Private m_stpLastPosition               As ScrollPositions
Private m_bolInDrag                     As Boolean
Private m_XDrag                         As Long
Private m_YDrag                         As Long
Private m_bolPanActive                  As Boolean

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Paint()
Event Resize()
Event Scroll()
Event HScroll(ScrValue As Integer)
Event VScroll(ScrValue As Integer)
Event ZoomChanged(ByVal ZoomPercent As Long)
Event ZoomInClick()
Event ZoomOutClick()
'Default Property Values:
'Const m_def_H_Scroll = 0
'Const m_def_V_Scroll = 0
Const m_def_AutoSize = False
'Property Variables:
'Dim m_H_Scroll As Long
'Dim m_V_Scroll As Long
Dim m_AutoSize As Boolean


Public Property Get AllowZoomOut() As Boolean
    AllowZoomOut = m_bolAllowOut
End Property

Public Property Let AllowZoomOut(Value As Boolean)
    m_bolAllowOut = Value
    PropertyChanged "AllowZoomOut"
End Property

Public Property Get AllowZoomIn() As Boolean
    AllowZoomIn = m_bolAllowIn
End Property

Public Property Let AllowZoomIn(Value As Boolean)
    m_bolAllowIn = Value
    PropertyChanged "AllowZoomIn"
End Property

Public Property Get AutoRedraw() As Boolean
    AutoRedraw = picView.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    picView.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = picView.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picBuffer.BackColor = New_BackColor
    picMain.BackColor = New_BackColor
    picTemp.BackColor = New_BackColor
    picView.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get HasDC() As Boolean
    HasDC = picView.HasDC
End Property

Public Property Get hDC() As Long
    hDC = picView.hDC
End Property

Public Property Get hwnd() As Long
    hwnd = picView.hwnd
End Property

Public Property Get Image() As Picture
    Set Image = picView.Image
End Property

Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get PanActive() As Boolean
    PanActive = m_bolPanActive
End Property

Public Property Let PanActive(ByVal Value As Boolean)
    
' Fail through on local error.
'
On Error Resume Next

   ' Set the current value.
    m_bolPanActive = Value
    PropertyChanged "PanActive"
    
   ' Determine whether or not to load the cursor from the resource file.
    If Value = True Then
        picView.MouseIcon = Image2.Picture 'LoadResPicture(102, vbResCursor)
        picView.MousePointer = vbCustom
    Else
        picView.MousePointer = vbDefault
    End If

End Property

Public Property Get Picture() As StdPicture
    Set Picture = picBuffer.Picture
End Property

Public Property Set Picture(Value As StdPicture)
    
   ' Set the image to the buffer.
    Set picBuffer.Picture = Value
    
   ' Display the image.
    Call ShowPicture
    
   ' Store the original size for scaling purposes.
    m_udtOriginal.Height = picView.Height
    m_udtOriginal.Width = picView.Width
  
   ' Reflect the fact that the zoom is at 100%
    Zoom = 100
    
End Property

Public Property Get UseQuickBar() As Boolean
    UseQuickBar = m_bolUseQuickBar
End Property

Public Property Let UseQuickBar(Value As Boolean)
    m_bolUseQuickBar = Value
    PropertyChanged "UseQuickBar"
End Property

Public Property Get Zoom() As Double
    Zoom = m_dblPercentage
End Property

Public Property Let Zoom(New_Zoom As Double)
    m_dblPercentage = New_Zoom
    
   ' Match the image to the new value.
    SetZoom
    
    RaiseEvent ZoomChanged(New_Zoom)
End Property

Public Sub Cls()
    picView.Cls
End Sub

'**********************************************************************************************************************'
' Procedure : LoadImage
' Date      : 01.27.2004
' Purpose   : Provide a LoadImage method to the UserControl.  Method allows for direct path manipulation with a
'             supplied filepath, or runs interactively with no supplied filepath.
' Input     : FilePath (String)
' Output    : None
'**********************************************************************************************************************'
Public Sub LoadImage(Optional FilePath As String = vbNullString)
  
' Fail through on error.
On Error Resume Next

   ' Determine if the mode is static or interactive.
    If FilePath <> vbNullString Then
  
       ' Make sure that a non-empty path was provided.
        If FilePath = vbNullString Then
            Exit Sub
        End If
    
       ' Store the path.
        m_strFileName = FilePath
        
       ' Load the image into the buffer.
        picBuffer.Picture = LoadPicture(FilePath)
        
    End If
    
   ' Display the image.
    Call ShowPicture
    
   ' Store the original size for scaling purposes.
    m_udtOriginal.Height = picView.Height
    m_udtOriginal.Width = picView.Width
  
   ' Reflect the fact that the zoom is at 100%
    Zoom = 100
  
End Sub

Public Sub PaintPicture( _
       ByVal Picture As Picture, _
       ByVal X1 As Single, _
       ByVal Y1 As Single, _
       Optional ByVal Width1 As Variant, _
       Optional ByVal Height1 As Variant, _
       Optional ByVal X2 As Variant, _
       Optional ByVal Y2 As Variant, _
       Optional ByVal Width2 As Variant, _
       Optional ByVal Height2 As Variant, _
       Optional ByVal Opcode As Variant)
    
' Fail through on local error.
On Error Resume Next

   ' Relay method.
   If Picture.Handle <> 0 Then
    picView.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, vbSrcCopy
   ' Set picBuffer.Picture = Picture 'picView.Image

    'picView.Refresh
    'ShowPicture
    End If
'    picBuffer.Refresh
End Sub

'**********************************************************************************************************************'
' Procedure : RecallPosition
' Date      : 02.05.2004
' Purpose   : Revert to the last stored position, if the scrollbar sizes are the same.
' Input     : None
' Output    : None
'**********************************************************************************************************************'

Public Sub RecallPosition()

   ' Make sure both scrollbars match (prevents it from reseting the scrollbars on every resize).
    If hscPicture.Max = m_stpLastPosition.HorizontalScrollMax And _
       vscPicture.Max = m_stpLastPosition.VerticalScrollMax Then
       
       ' Reset to the last value.
        hscPicture.Value = m_stpLastPosition.HorizontalScrollPosition
        vscPicture.Value = m_stpLastPosition.VerticalScrollPosition
        
    End If
    
End Sub

'**********************************************************************************************************************'
' Procedure : StorePosition
' Date      : 02.05.2004
' Purpose   : Save the current scroll information
' Input     : None
' Output    : None
'**********************************************************************************************************************'

Public Sub StorePosition()

   ' Store it like it is.
    With m_stpLastPosition
        .HorizontalScrollMax = hscPicture.Max
        .HorizontalScrollPosition = hscPicture.Value
        .VerticalScrollMax = vscPicture.Max
        .VerticalScrollPosition = vscPicture.Value
    End With
    
End Sub

'**********************************************************************************************************************'
' Procedure : UnloadImage
' Date      : 02.04.2004
' Purpose   : Return the control to the state that it was in prior to loading an image.
' Input     : None
' Output    : None
'**********************************************************************************************************************'
Public Sub UnloadImage()

' Fail through on local error.
'
On Error Resume Next

   ' Wipe the member variables to their initial states.
    m_strFileName = vbNullString
    m_udtOriginal.Height = 0
    m_udtOriginal.Width = 0
    
   ' Clear the image from the picture properties that might be holding a reference.
    picBuffer.Picture = Nothing
    picView.Picture = Nothing
    
   ' Redraw the display area.
    picView.Cls
    
   ' Turn off the zoom buttons.
    tbrQuickFloat.Buttons(1).Enabled = False
    tbrQuickFloat.Buttons(2).Enabled = False
    
End Sub

Private Sub hscPicture_Change()
  
' Fail through on local error.
On Error Resume Next

   ' Match the picView's horizontal orientation with the scroll value.
    picView.Left = hscPicture.Value
    RaiseEvent HScroll(hscPicture.Value)
End Sub

Private Sub hscPicture_Scroll()
    
' Fail through on local error.
On Error Resume Next

   ' Relay event.
    RaiseEvent Scroll
    
End Sub

Private Sub picMain_DblClick()

' Fail through on local error.
On Error Resume Next

   ' Relay event.
    RaiseEvent DblClick
    
End Sub

Private Sub picMain_Resize()

' Fail through on local error.
On Error Resume Next
'
'   ' Set the orientation of the horizontal scrollbar.
'    With hscPicture
'
'        .Left = 0
'
'        If vscPicture.Visible Then
'            .Top = ScaleY(UserControl.Height, vbTwips, vbPixels) - hscPicture.Height - GetSystemMetrics(SM_CYDLGFRAME)
'            .Width = ScaleX(UserControl.Width, vbTwips, vbPixels) - vscPicture.Width - _
'                     (2 * GetSystemMetrics(SM_CXBORDER))
'        Else
'            .Top = ScaleY(UserControl.Height, vbTwips, vbPixels) - hscPicture.Height - GetSystemMetrics(SM_CYDLGFRAME)
'            .Width = ScaleX(UserControl.Width, vbTwips, vbPixels) - GetSystemMetrics(SM_CXDLGFRAME)
'        End If
'    End With
'
'   ' Set the orientation of the vertical scrollbar.
'    With vscPicture
'
'        .Top = 0
'
'        If hscPicture.Visible Then
'            .Left = ScaleX(UserControl.Width, vbTwips, vbPixels) - vscPicture.Width - GetSystemMetrics(SM_CXDLGFRAME)
'            .Height = ScaleY(UserControl.Height, vbTwips, vbPixels) - hscPicture.Height - _
'                      (2 * GetSystemMetrics(SM_CYBORDER))
'        Else
'            .Left = ScaleX(UserControl.Width, vbTwips, vbPixels) - vscPicture.Width - GetSystemMetrics(SM_CXDLGFRAME)
'            .Height = ScaleY(UserControl.Height, vbTwips, vbPixels) - GetSystemMetrics(SM_CYDLGFRAME)
'        End If
'    End With
    
   ' Move the generic blocking box to the bottom right corner.
    picTemp.Move hscPicture.Width, vscPicture.Height, vscPicture.Width, hscPicture.Height
    
   ' Make sure that the scrollbars are all good.
    CheckForScrolls
    
End Sub

Private Sub picView_Click()
    RaiseEvent Click
End Sub

Private Sub picView_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picView_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picView_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub


Private Sub picView_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Private Sub picView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
Dim rctCurrentMain                      As RECT
Dim pntCursor                           As POINTAPI

' Fail through on local errors.
On Error Resume Next

    RaiseEvent MouseDown(Button, Shift, X, Y)
    
   ' Check to see if panning is active.
    If m_bolPanActive Then
       ' Panning can only be active if one or more scroll bars are visible and the left button is down.
        m_bolInDrag = ((hscPicture.Visible Or vscPicture.Visible) And (Button = vbLeftButton))
       ' Verify.
        If m_bolInDrag Then
        
           ' Get the mouse and picturebox positions.  These will be used for delta calculations.
            Call GetWindowRect(picMain.hwnd, rctCurrentMain)
            Call GetCursorPos(pntCursor)
            
           ' Establish a baseling for future deltas.
            m_XDrag = pntCursor.X - rctCurrentMain.Left
            m_YDrag = pntCursor.Y - rctCurrentMain.Top
            
           ' Load the pan icon from the resource file.
            picView.MouseIcon = Image1.Picture 'LoadResPicture(101, vbResCursor)
            picView.MousePointer = vbCustom
            
        End If
    
    End If
    
End Sub

Private Sub picView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    

Dim lngCurrentX                         As Long
Dim lngCurrentY                         As Long
Dim rctCurrentMain                      As RECT
Dim pntCursor                           As POINTAPI

' Fail through on local errors.
On Error Resume Next

   ' Make sure that it is being dragged.
    If m_bolInDrag Then
        
       ' Get the position for the delta.
        Call GetWindowRect(picMain.hwnd, rctCurrentMain)
        Call GetCursorPos(pntCursor)
        
       ' Calcualte current logical position.
        lngCurrentX = pntCursor.X - rctCurrentMain.Left
        lngCurrentY = pntCursor.Y - rctCurrentMain.Top
                
       ' Check to see if the scroll is visible and available for movement.
        If hscPicture.Visible = True Then
            If lngCurrentX < m_XDrag Then
                hscPicture.Value = hscPicture.Value + (Abs(m_XDrag - lngCurrentX))
            ElseIf lngCurrentX > m_XDrag Then
                hscPicture.Value = hscPicture.Value - (Abs(lngCurrentX - m_XDrag))
            End If
        End If
        
       ' Check to see if the scroll is visible and available for movement.
        If vscPicture.Visible = True Then
            If lngCurrentY < m_YDrag Then
                vscPicture.Value = vscPicture.Value + (Abs(m_YDrag - lngCurrentY))
            ElseIf lngCurrentY > m_YDrag Then
                vscPicture.Value = vscPicture.Value - (Abs(lngCurrentY - m_YDrag))
            End If
        End If
                
       ' Store the new position for the next delta.
        m_YDrag = lngCurrentY
        m_XDrag = lngCurrentX
        
    End If
    
End Sub

Private Sub picView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
' Fail through on local error.
On Error Resume Next

    RaiseEvent MouseUp(Button, Shift, X, Y)
    
   ' Not draggin anymore
    m_bolInDrag = False

   ' Make sure that the panning is active.
    If m_bolPanActive Then
    
       ' Restore the cursor the the standard pan.
        picView.MouseIcon = Image2.Picture 'LoadResPicture(102, vbResCursor)
        picView.MousePointer = vbCustom
        
    Else
    
       ' Use the default pointer.
        picView.MousePointer = vbDefault
        
    End If
    
End Sub

Private Sub picView_Paint()
    RaiseEvent Paint
End Sub

Private Sub tbrQuickFloat_ButtonClick( _
        ByVal Button As MSComctlLib.Button)

' Fail through on local error.
'
On Error Resume Next

   ' Determine whether to raise a ZoomInClick(1) or ZoomOutClick(2)
   '
    If Button.Index = 1 Then
        RaiseEvent ZoomInClick
    Else
        RaiseEvent ZoomOutClick
    End If
    
End Sub

Private Sub timToolbar_Timer()

' Fail through on local error.
'
On Error Resume Next

Dim lnghWnd                             As Long
Dim lngResult                           As Long
Dim pntMouse                            As POINTAPI
Dim tbtButtonInfo                       As TBBUTTON
    
   ' Bail if the user is not going to be using the QuickBar
    If Not m_bolUseQuickBar Then
        Exit Sub
    End If
    
   ' Determine visibility of the zoom in button.
    If Not (tbrQuickFloat.Buttons(1).Enabled = m_bolAllowIn) Then
        tbrQuickFloat.Buttons(1).Enabled = m_bolAllowIn
    End If
    
   ' Determine visibility of the zoom out button.
    If Not (tbrQuickFloat.Buttons(2).Enabled = m_bolAllowOut) Then
        tbrQuickFloat.Buttons(2).Enabled = m_bolAllowOut
    End If
    
   ' Make sure that an image exists by checking for the image handle.
    If picBuffer.Picture.Handle = 0 Then
        
       ' Bounce!
        Exit Sub
    
    End If
    
   ' Populate the point structure from the current mouse position.
    Call GetCursorPos(pntMouse)

   ' Retrieve the handle of the current window described by the point structure.
    lnghWnd = WindowFromPoint(pntMouse.X, pntMouse.Y)
    
   ' Populate the button info.
    lngResult = SendMessage(lnghWnd, TB_GETBUTTON, 0, tbtButtonInfo)
    
   ' Evaluate the current mouse position against the controls inside of the usercontrol.  Note the last
   ' clause which checks for any handle that is two levels deep inside of the usercontrol that is recognized
   ' as a button.
    If lnghWnd = UserControl.hwnd Or _
       lnghWnd = picMain.hwnd Or _
       lnghWnd = tbrQuickFloat.hwnd Or _
       lnghWnd = picTemp.hwnd Or _
       lnghWnd = picBuffer.hwnd Or _
       lnghWnd = picView.hwnd Or _
       (tbtButtonInfo.idCommand > 0 And GetParent(GetParent(lnghWnd)) = UserControl.hwnd) Then

       ' Make sure that it is visible.
        tbrQuickFloat.Visible = True

    Else

       ' Make sure that it is not visible.
        tbrQuickFloat.Visible = False
        
    End If
        
End Sub

Private Sub UserControl_Initialize()

    picTemp.ZOrder
    
End Sub

Private Sub UserControl_InitProperties()

' Fail through on local error.
On Error Resume Next

    timToolbar.Enabled = UserControl.Ambient.UserMode
    
    If timToolbar.Enabled Then
    
       ' Set the handle to subclass.
        g_lngTargetHwnd = hscPicture.Parent.hwnd
        
       ' ...line, and sinker.
        modScrollFix.Hook
        
    End If
    m_AutoSize = m_def_AutoSize
'    m_H_Scroll = m_def_H_Scroll
'    m_V_Scroll = m_def_V_Scroll
End Sub

Private Sub Usercontrol_KeyDown(KeyCode As Integer, Shift As Integer)
        
' Fail through on local error.
On Error Resume Next

   ' Look for the left and right movement.
    If (((KeyCode = 39) Or (KeyCode = 37)) And (hscPicture.Value)) Then
        hscPicture.SetFocus
        Exit Sub
    End If
    
   ' Look for the up and down movement.
    If (((KeyCode = 38) Or (KeyCode = 40)) And (vscPicture.Value)) Then
        vscPicture.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

' Fail through on local error.
On Error Resume Next

   ' Make sure that the timer is not running if it is in design mode.
    timToolbar.Enabled = UserControl.Ambient.UserMode

    If timToolbar.Enabled Then
    
       ' Set the handle to subclass.
        g_lngTargetHwnd = picMain.hwnd
        
       ' ...line, and sinker.
        modScrollFix.Hook
        
    End If
    
   ' Populate the value from the property bag.
    picView.AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
    picView.BackColor = PropBag.ReadProperty("BackColor", &H8000000C)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    AllowZoomIn = PropBag.ReadProperty("AllowZoomIn", False)
    AllowZoomOut = PropBag.ReadProperty("AllowZoomOut", False)
    UseQuickBar = PropBag.ReadProperty("UseQuickBar", False)
    PanActive = PropBag.ReadProperty("PanActive", False)
    picView.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 65)
    picView.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 129)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
   ' hscPicture.Value = PropBag.ReadProperty("H_Scroll", 0)
  '  vscPicture.Value = PropBag.ReadProperty("V_Scroll", 0)
End Sub

Private Sub UserControl_Resize()

' Fail through on local error.
On Error Resume Next

   ' Move the main picturebox to match the full size of the usercontrol.  The remainder of the resizing will be
   ' accomplished through the resize events of the children.
    picMain.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    
   ' Make sure that the picturebox is in front to block the corner.
    picTemp.ZOrder
    
   ' Raise it on down the line...
    RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
         CheckForScrolls
End Sub

Private Sub UserControl_Terminate()
    
' Fail through on local error.
On Error Resume Next

   ' Check to make sure that they control is not firing this event in the IDE.
    If UserControl.Ambient.UserMode = True Then
        
       ' Terminate Subsclassing.
        modScrollFix.Unhook
        
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

' Fail through on local error.
On Error Resume Next

   ' Store to the property bag.
    Call PropBag.WriteProperty("AutoRedraw", picView.AutoRedraw, True)
    Call PropBag.WriteProperty("BackColor", picView.BackColor, &H8000000C)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("AllowZoomIn", m_bolAllowIn, False)
    Call PropBag.WriteProperty("AllowZoomOut", m_bolAllowOut, False)
    Call PropBag.WriteProperty("UseQuickBar", m_bolUseQuickBar, False)
    Call PropBag.WriteProperty("PanActive", m_bolPanActive, False)
    Call PropBag.WriteProperty("ScaleHeight", picView.ScaleHeight, 65)
    Call PropBag.WriteProperty("ScaleWidth", picView.ScaleWidth, 129)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
   ' Call PropBag.WriteProperty("H_Scroll", hscPicture.Value, 0)
   ' Call PropBag.WriteProperty("V_Scroll", vscPicture.Value, 0)
End Sub
Private Sub vscPicture_Change()
  
' Fail through on local error.
On Error Resume Next

   ' Synch them.
    picView.Top = vscPicture.Value
    RaiseEvent VScroll(vscPicture.Value)
End Sub

Private Sub CheckForScrolls()
    Dim hv As Boolean, vv As Boolean
' Fail through on local error.
On Error Resume Next

   ' Check to see if the width of the image requires the use of the scrollbars.
    If (picView.Width < picMain.Width - GetSystemMetrics(SM_CXHTHUMB)) Then
        hscPicture.Value = hscPicture.Min
        hscPicture.Visible = False
        picView.Left = (picMain.Width - picView.Width) / 2
        hv = False
    Else
        With hscPicture
            .Visible = True
            .ZOrder
            .Min = 0
            .Max = -(picView.Width - (picMain.Width - GetSystemMetrics(SM_CXHTHUMB)) + 4)
            '.Value = (.Max - .Min) / 2
        End With
        hv = True
    End If
    
   ' Check to see if the height of the image requires the use of the scrollbars.
    If (picView.Height < picMain.Height - GetSystemMetrics(SM_CYVTHUMB)) Then
        vscPicture.Value = vscPicture.Min
        vscPicture.Visible = False
        picView.Top = (picMain.Height - picView.Height) / 2
        vv = False
    Else
        With vscPicture
            .Visible = True
            .ZOrder
            .Min = 0
            .Max = -(picView.Height - (picMain.Height - GetSystemMetrics(SM_CYVTHUMB)) + 4)
            '.Value = (.Max - .Min) / 2
        End With
        vv = True
    End If

   ' Make sure the space filler is not visible if one of the scrollbars is not visible.
    'picTemp.Visible = (hscPicture.Visible And vscPicture.Visible)
    picTemp.Visible = (vv And hv)
   ' Move the generic blocking box to the bottom right corner.
    picTemp.Move hscPicture.Width, vscPicture.Height, vscPicture.Width, hscPicture.Height
    picTemp.ZOrder
    
   ' Set the orientation of the horizontal scrollbar.
    With hscPicture
    
       ' Lock the origin.
        .Left = 0
        
        If (picView.Height < picMain.Height - GetSystemMetrics(SM_CYVTHUMB)) Then
            .Top = ScaleY(UserControl.Height, vbTwips, vbPixels) - hscPicture.Height - GetSystemMetrics(SM_CYDLGFRAME)
            .Width = ScaleX(UserControl.Width, vbTwips, vbPixels) - GetSystemMetrics(SM_CXDLGFRAME)
        Else
            .Top = ScaleY(UserControl.Height, vbTwips, vbPixels) - hscPicture.Height - GetSystemMetrics(SM_CYDLGFRAME)
            .Width = ScaleX(UserControl.Width, vbTwips, vbPixels) - vscPicture.Width - (2 * GetSystemMetrics(SM_CXBORDER))
        End If
    End With
    
   ' Set the orientation of the vertical scrollbar.
    With vscPicture
        
       ' Lock the origin.
        .Top = 0
        
       ' Make sure the size is correct.
       If (picView.Width < picMain.Width - GetSystemMetrics(SM_CXHTHUMB)) Then
          .Left = ScaleX(UserControl.Width, vbTwips, vbPixels) - vscPicture.Width - GetSystemMetrics(SM_CXDLGFRAME)
          .Height = ScaleY(UserControl.Height, vbTwips, vbPixels) - GetSystemMetrics(SM_CYDLGFRAME)
       Else
          .Left = ScaleX(UserControl.Width, vbTwips, vbPixels) - vscPicture.Width - GetSystemMetrics(SM_CXDLGFRAME)
          .Height = ScaleY(UserControl.Height, vbTwips, vbPixels) - hscPicture.Height - (2 * GetSystemMetrics(SM_CYBORDER))
       End If
    End With
    
End Sub

'**********************************************************************************************************************'
' Procedure : SetZoom
' Date      : 02.04.2004
' Purpose   : Synch the member zoom value with the display
' Input     : None
' Output    : None
'**********************************************************************************************************************'
Private Sub SetZoom()

' Fail through on local errors.
On Error Resume Next

   ' Make sure that there is a filename and the there is an image in the buffer.
    If m_strFileName = vbNullString And picBuffer.Picture.Handle = 0 Then
        Exit Sub
    End If
    
   ' Lock the window during the resize and paint.
    LockWindowUpdate picView.hwnd
    
   ' Resize the view.
    picView.Width = m_udtOriginal.Width * (m_dblPercentage / 100)
    picView.Height = m_udtOriginal.Height * (m_dblPercentage / 100)
    
   ' Paint the buffer onto the view.
    Call picView.PaintPicture(picBuffer.Picture, 0, 0, picView.Width, picView.Height)

   ' Make sure that the scroll bars reflect the current view.
    Call CheckForScrolls
    
   ' Allow for redraw.
    LockWindowUpdate 0
    
End Sub

'**********************************************************************************************************************'
' Procedure : ShowPicture
' Date      : 02.04.2004
' Purpose   : Display the buffered picture.
' Input     : None
' Output    : None
'**********************************************************************************************************************'
Private Sub ShowPicture()

' Fail through on local error.
On Error Resume Next

   ' Make sure that the view is clean and visible.
    picView.Visible = True
    picView.Cls
    
   ' Load the picture directly into the view from the buffer
    Set picView.Picture = picBuffer.Picture
    
   ' Center the view in the main picture box.
    If picView.Width < picMain.Width Then
        picView.Left = (picMain.Width - picView.Width - GetSystemMetrics(SM_CYVTHUMB)) / 2
    Else
        picView.Left = (picMain.Width - picView.Width) / 2
    End If
    
    If picView.Height < picMain.Height Then
        picView.Top = (picMain.Height - picView.Height - GetSystemMetrics(SM_CYVTHUMB)) / 2
    Else
        picView.Top = (picMain.Height - picView.Height) / 2
    End If
            
   ' Make sure that the scroll bars reflect the current view.
    Call CheckForScrolls
  
    DoEvents
        
   ' Rerun the scroll check.  For some reason that I have not been able to track down, picTemp will sometimes
   ' disappear until the form is resized unless this is run again.
    Call CheckForScrolls

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picView,picView,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = picView.ScaleHeight
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picView,picView,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = picView.ScaleWidth
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
              
   ' Display the image.
    Call ShowPicture
    
   ' Store the original size for scaling purposes.
    m_udtOriginal.Height = picView.Height
    m_udtOriginal.Width = picView.Width
  
   ' Reflect the fact that the zoom is at 100%
    Zoom = 100
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    Dim zx As Long, zy As Long, i As Integer
    m_AutoSize = New_AutoSize
    If New_AutoSize Then
      If picBuffer.Picture <> 0 Then
         If vscPicture.Visible = True Or hscPicture.Visible = True Then
         For i = 1 To 1000
         Zoom = Zoom - 10
           If vscPicture.Visible = False And hscPicture.Visible = False Then Exit For
         Next
         End If
         picView.Cls
         Call picView.PaintPicture(picBuffer.Picture, 0, 0, picView.Width, picView.Height)
      End If
    End If
    PropertyChanged "AutoSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=hscPicture,hscPicture,-1,Value
Public Property Get H_Scroll() As Integer
Attribute H_Scroll.VB_Description = "Returns/sets the value of an object."
    H_Scroll = hscPicture.Value
End Property

Public Property Let H_Scroll(ByVal New_H_Scroll As Integer)
    hscPicture.Value() = New_H_Scroll
    PropertyChanged "H_Scroll"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=vscPicture,vscPicture,-1,Value
Public Property Get V_Scroll() As Integer
Attribute V_Scroll.VB_Description = "Returns/sets the value of an object."
    V_Scroll = vscPicture.Value
End Property

Public Property Let V_Scroll(ByVal New_V_Scroll As Integer)
    vscPicture.Value() = New_V_Scroll
    PropertyChanged "V_Scroll"
End Property

