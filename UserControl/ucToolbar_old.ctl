VERSION 5.00
Begin VB.UserControl ucToolbar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   ControlContainer=   -1  'True
   ForeColor       =   &H80000014&
   LockControls    =   -1  'True
   ScaleHeight     =   75
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
   Begin VB.Timer tmrTip 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblTipRect 
      BackStyle       =   0  'Transparent
      Height          =   270
      Left            =   -375
      TabIndex        =   0
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "ucToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucToolbar.ctl (2.1)
' Author:        Carles P.V.
' Dependencies:  -
' First release: 2003.09.16
'================================================
'
' LOG:
'
' - 2003.10.12: Fixed <ButtonCheck> event raising.
' - 2003.10.12: Improved <Over> effect (on Check and Option buttons)
' - 2003.10.12: Added Backcolor property (3D-colors automaticaly generated)
' - 2003.10.13: Fixed! <ButtonCheck> event raising.
' - 2003.10.15: Fixed state <release> updating of Normal button on
'               CANCEL_MODE (Ctrl+Esc).
' - 2003.10.15: Fixed state <release> updating of Normal button on
' - 2003.10.18: Sub <BuildToolbar> to Function <(boolean)BuildToolbar>
' - 2003.10.18: Added (sub)SetTooltip[button]/(function)GetTooltip[button]
' - 2003.10.20: <ButtonCheck> (Check buttons) is also raised on unchecking.
' - 2003.10.30: Improved 'disabled icon painting' for custom colors.
'               Thanks to LaVolpe for this nice tip. Original post at:
' - 2003.11.02: <ButtonClick> also raised with Option/Check buttons
' - 2003.11.02: Added 'Check' param. to <CheckButton> sub. (Thanks to VJA).
' - 2009.03.19: Change view disabled button.
Option Explicit

'-- API:

Private Type RECT2
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const ILC_MASK        As Long = &H1
Private Const ILC_COLORDDB    As Long = &HFE
Private Const ILD_TRANSPARENT As Long = 1
Private Const DST_ICON        As Long = &H3
Private Const DSS_UNION       As Long = &H10
Private Const DSS_DISABLED As Long = &H20&
Private Const DSS_MONO        As Long = &H80
Private Const CLR_INVALID     As Long = &HFFFF
Private Const PS_SOLID        As Long = 0

Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT2, ByVal dX As Long, ByVal dy As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT2, ByVal m_hCheckBrush As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function ImageList_Create Lib "comctl32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_AddMasked Lib "comctl32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32" (ByVal hIml As Long, ByVal I As Long, ByVal hDCDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_GetImageCount Lib "comctl32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const VK_LBUTTON As Long = &H1
Private Const VK_RBUTTON As Long = &H2
Private Const VK_MBUTTON As Long = &H4

'//

'-- Public Enums.:
Public Enum tbOrientationConstants
    [tbHorizontal] = 0
    [tbVertical] = 1
End Enum

'-- Private Enums.:
Private Enum tbButtonStateConstants
    [btDown] = -1
    [btFlat] = 0
    [btOver] = 1
End Enum

Private Enum tbButtonTypeConstants
    [btNormal] = 0
    [btCheck] = 1
    [btOption] = 2
End Enum

Private Enum tbMouseEventConstants
    [btMouseDown] = -1
    [btMouseMove] = 0
    [btMouseUp] = 1
End Enum

'-- Private Types:
Private Type tButton
    Type      As tbButtonTypeConstants
    STATE     As tbButtonStateConstants
    Enabled   As Boolean
    Checked   As Boolean
    Over      As Boolean
    Separator As RECT2
    Tooltip   As String
End Type

'-- Private Constants:
Private Const BT_STNORMAL  As String = "N"
Private Const BT_STCHECK   As String = "C"
Private Const BT_STOPTION  As String = "O"
Private Const BT_SEPARATOR As String = "|"

'-- Default Property Values:
Private Const m_def_BarOrientation As Integer = [tbHorizontal]
Private Const m_def_BarEdge        As Boolean = 0

'-- Property Variables:
Private m_BarOrientation  As tbOrientationConstants
Private m_BarEdge         As Boolean

'-- Private Variables:
Private m_hIL             As Long    ' Image list handle
Private m_hCheckBrush     As Long    ' Brush (check effect)
Private m_hFaceBrush      As Long    ' Brush (button face)
Private m_hHighlightBrush As Long    ' Brush (button highlight for disabled icon)
Private m_hShadowBrush    As Long    ' Brush (button shadow for disabled icon)
Private m_hHighlightPen   As Long    ' Pen (button highlight)
Private m_hShadowPen      As Long    ' Pen (button shadow)
Private m_BarRect         As RECT2   ' Bar rectangle
Private m_ExtRect()       As RECT2   ' Button rects. (edge area)
Private m_ClkRect()       As RECT2   ' Button rects. (click area)
Private m_uButton()       As tButton ' Buttons
Private m_Tooltip()       As String  ' Tool tips
Private m_Count           As Integer ' Button count
Private m_LastOver        As Integer ' Last over
Private m_IconSize        As Integer ' Icon size (W = H)
Private m_ButtonSize      As Integer ' Button size (W = H)
Private m_Mouse           As Integer ' Temp. mouse button

'-- Event Declarations:
Public Event ButtonClick(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xLeft As Long, ByVal yTop As Long)
Public Event ButtonCheck(ByVal Index As Long, ByVal xLeft As Long, ByVal yTop As Long)

'==================================================================================================
' UserControl
'==================================================================================================

Private Sub UserControl_Initialize()

  Dim aIdx           As Byte
  Dim nBytes(1 To 8) As Integer
  Dim hBitmap        As Long
    
    '-- Build brush for check effect
    For aIdx = 1 To 8 Step 2
        nBytes(aIdx) = &HAA
        nBytes(aIdx + 1) = &H55
    Next aIdx
    hBitmap = CreateBitmap(8, 8, 1, 1, nBytes(1))
    m_hCheckBrush = CreatePatternBrush(hBitmap)
    DeleteObject hBitmap
End Sub

Private Sub UserControl_Terminate()

    '-- Destroy image list, ...
    Call pvDestroyIL
    '   ... pens and brushes
    If (m_hCheckBrush) Then DeleteObject m_hCheckBrush
    If (m_hFaceBrush) Then DeleteObject m_hFaceBrush
    If (m_hHighlightPen) Then DeleteObject m_hHighlightPen
    If (m_hShadowPen) Then DeleteObject m_hShadowPen
    If (m_hHighlightBrush) Then DeleteObject m_hHighlightBrush
    If (m_hShadowBrush) Then DeleteObject m_hShadowBrush
End Sub

'//

Private Sub UserControl_Show()
    '-- Refresh on start up
    Call pvRefresh
End Sub

Private Sub UserControl_Resize()
    
    '-- Adjust for alignment
    Select Case m_BarOrientation
        Case [tbHorizontal]
            m_BarRect.X2 = ScaleWidth
        Case [tbVertical]
            m_BarRect.Y2 = ScaleHeight
    End Select
    '-- Refresh whole control
    FillRect hdc, m_BarRect, m_hFaceBrush
    Call pvRefresh
End Sub

'//

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  Dim nBtn As Integer
    
    '-- Restore last
    If (m_LastOver) Then
        pvUpdateButtonState m_LastOver, 0, 0, [btMouseMove]
    End If
    '-- Update tooltip label pos.
    For nBtn = 1 To m_Count
        If (PtInRect(m_ExtRect(nBtn), X, Y) And m_uButton(nBtn).Enabled) Then
            Call pvSetTipArea(nBtn)
            m_LastOver = nBtn
        End If
    Next nBtn
End Sub

Private Sub lblTipRect_DblClick()
     
    If (GetAsyncKeyState(VK_RBUTTON) >= 0 And GetAsyncKeyState(VK_MBUTTON) >= 0) Then '*
        '-- Preserve second click
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    End If
'*: Should be previously checked GetSystemMetrics(SM_SWAPBUTTON)
End Sub

Private Sub lblTipRect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (m_LastOver) Then
        If (m_uButton(m_LastOver).Enabled And m_Mouse = vbEmpty) Then
            '-- Refresh state
            Call pvUpdateButtonState(m_LastOver, -1, Button, [btMouseDown])
        End If
        m_Mouse = Button
    End If
End Sub

Private Sub lblTipRect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  Dim lx As Long
  Dim ly As Long
    
    If (m_LastOver) Then
        If (m_uButton(m_LastOver).Enabled) Then
            '-- Translate to [pixels]
            lx = X \ Screen.TwipsPerPixelX + lblTipRect.Left
            ly = Y \ Screen.TwipsPerPixelY + lblTipRect.Top
            '-- Refresh state
            Call pvUpdateButtonState(m_LastOver, PtInRect(m_ExtRect(m_LastOver), lx, ly) <> 0, Button, [btMouseMove])
        End If
        If (Button = vbLeftButton) Then tmrTip.Enabled = -1
    End If
End Sub

Private Sub lblTipRect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  Dim lx As Long
  Dim ly As Long
    
    If (m_LastOver) Then
        If (m_uButton(m_LastOver).Enabled) Then
            '-- Translate to [pixels]
            lx = X \ Screen.TwipsPerPixelX + lblTipRect.Left
            ly = Y \ Screen.TwipsPerPixelY + lblTipRect.Top
            '-- Refresh state
            Call pvUpdateButtonState(m_LastOver, PtInRect(m_ExtRect(m_LastOver), lx, ly) <> 0, Button, [btMouseUp])
            m_Mouse = 0: ReleaseCapture
        End If
    End If
End Sub

'==================================================================================================
' Methods
'==================================================================================================

Public Sub Refresh()
    '-- Refresh whole bar
    Call pvRefresh
End Sub

Public Function BuildToolbar(Image As StdPicture, _
                             ByVal MaskColor As OLE_COLOR, _
                             ByVal IconSize As Integer, _
                             Optional ByVal FormatMask As String) As Boolean
  Dim nIdx As Integer
  Dim nBtn As Integer
  Dim sKey As String
  Dim lPos As Long
    
    If (pvExtractImages(Image, MaskColor, IIf(IconSize > 0, IconSize, 1))) Then
        
        '-- Missing 'FormatMask': Normal buttons, no separators
        If (FormatMask = vbNullString) Then
            FormatMask = String$(ImageList_GetImageCount(m_hIL), BT_STNORMAL)
        End If
        
        '-- Button ext. size (image[state] and edge offsets)
        m_ButtonSize = m_IconSize + 7
        
        '-- Extract buttons...
        Do While nIdx < Len(FormatMask)
            
            '-- Key count / extract key
            nIdx = nIdx + 1
            sKey = UCase$(mid$(FormatMask, nIdx, 1))
            
            Select Case sKey
                
                '-- Normal button, check button and option buttons
                Case BT_STNORMAL, BT_STCHECK, BT_STOPTION
                
                    nBtn = nBtn + 1
                    lPos = lPos + m_ButtonSize
                    
                    '-- Redim. button rectangles
                    ReDim Preserve m_ExtRect(1 To nBtn)
                    ReDim Preserve m_ClkRect(1 To nBtn)
                    ReDim Preserve m_uButton(1 To nBtn)
                    '-- Store button type
                    Select Case sKey
                        Case BT_STNORMAL: m_uButton(nBtn).Type = [btNormal]
                        Case BT_STCHECK:  m_uButton(nBtn).Type = [btCheck]
                        Case BT_STOPTION: m_uButton(nBtn).Type = [btOption]
                    End Select
                    '-- Enabled [?]
                    m_uButton(nBtn).Enabled = UserControl.Enabled
                    
                    '-- Button ext. rect.
                    Select Case m_BarOrientation
                        Case [tbHorizontal]
                            SetRect m_ExtRect(nBtn), lPos - m_ButtonSize, 0, lPos, m_ButtonSize - 1
                        Case [tbVertical]
                            SetRect m_ExtRect(nBtn), 0, lPos - m_ButtonSize, m_ButtonSize - 1, lPos
                    End Select
                    OffsetRect m_ExtRect(nBtn), 1, 1
                    '-- Button click rect.
                    m_ClkRect(nBtn) = m_ExtRect(nBtn): InflateRect m_ClkRect(nBtn), -2, -2
               
                '-- Separator
                Case BT_SEPARATOR
                
                    lPos = lPos + 6
                    With m_ClkRect(nBtn)
                        Select Case m_BarOrientation
                            Case [tbHorizontal]
                                SetRect m_uButton(nBtn).Separator, .X2 + 4, .Y1, .X2 + 5, .Y2
                            Case [tbVertical]
                                SetRect m_uButton(nBtn).Separator, .X1, .Y2 + 4, .X2, .Y2 + 5
                        End Select
                    End With
            End Select
        Loop
        
        '-- Resize control
        With m_ExtRect(nBtn)
            UserControl.Width = (.X2 + 1) * Screen.TwipsPerPixelX
            UserControl.Height = (.Y2 + 1) * Screen.TwipsPerPixelY
        End With
        SetRect m_BarRect, 0, 0, ScaleWidth, ScaleHeight
        
        '-- Buttons count / success
        m_Count = nBtn
        BuildToolbar = (m_Count > 0)
    End If
End Function

Public Sub SetTooltips(ByVal TooltipsList As String)
    Dim Index As Long
    '-- Extract tooltips...
    m_Tooltip() = Split(TooltipsList, BT_SEPARATOR)
    For Index = 1 To m_Count
       m_uButton(Index).Tooltip = m_Tooltip(Index - 1)
    Next
    Erase m_Tooltip
End Sub

Public Sub SetTooltip(ByVal Index As Integer, ByVal Tooltip As String)
    m_uButton(Index).Tooltip = Tooltip
    'm_Tooltip(Index - 1) = Tooltip
End Sub

Public Function GetTooltip(ByVal Index As Integer) As String
    GetTooltip = m_uButton(Index).Tooltip
End Function

Public Sub EnableButton(ByVal Index As Integer, ByVal Enable As Boolean)
    Call pvEnableButton(Index, Enable)
End Sub
Public Function IsButtonEnabled(ByVal Index As Integer) As Boolean
    IsButtonEnabled = m_uButton(Index).Enabled
End Function

Public Sub CheckButton(ByVal Index As Integer, ByVal Check As Boolean)

    If (m_Count) Then
        If (Index And Index <= m_Count) Then
            If (m_uButton(Index).Type <> [btNormal] And m_uButton(Index).Checked <> Check) Then
                
                '-- Update button
                With m_uButton(Index)
                    .Checked = Check
                    .STATE = [btDown] And Check
                End With
                Call pvRefresh(Index)
                Call pvUpdateOptionButtons(Index)
                '-- Update Tooltip label pos.
                Call pvSetTipArea(Index)
                '-- Store <last over> index
                m_LastOver = Index
            
                '-- Raise <Check> event
                With m_ExtRect(Index)
                    RaiseEvent ButtonCheck(Index, .X1, .Y1)
                End With
            End If
        End If
    End If
End Sub
Public Function IsButtonChecked(ByVal Index As Integer) As Boolean
    IsButtonChecked = m_uButton(Index).Checked
End Function

'==================================================================================================
' Private
'==================================================================================================

Private Function pvExtractImages(Image As StdPicture, ByVal MaskColor As OLE_COLOR, ByVal IconSize As Integer) As Boolean
    
    '-- Extract images
    If (Not Image Is Nothing) Then
        If (pvCreateIL(IconSize)) Then
            pvExtractImages = (ImageList_AddMasked(m_hIL, Image.Handle, pvTranslateColor(MaskColor)) <> -1)
        End If
    End If
End Function

Private Function pvCreateIL(ByVal IconSize As Integer) As Boolean
     
    '-- Destroy previous [?]
    Call pvDestroyIL
    '-- Create the image list object:
    m_hIL = ImageList_Create(IconSize, IconSize, ILC_MASK Or ILC_COLORDDB, 0, 0)
    If (m_hIL <> 0) And (m_hIL <> -1) Then
        m_IconSize = IconSize
        pvCreateIL = -1
      Else
        m_hIL = 0
    End If
End Function

Private Sub pvDestroyIL()

    '-- Kill the image list if we have one:
    If (m_hIL <> 0) Then
        ImageList_Destroy m_hIL
        m_hIL = 0
    End If
End Sub

'//

Private Sub pvSetTipArea(ByVal Index As Integer)
    
    '-- Move label
    Select Case m_BarOrientation
        Case [tbHorizontal]
            lblTipRect.Move m_ExtRect(Index).X1, 0, m_ButtonSize, m_ButtonSize
        Case [tbVertical]
            lblTipRect.Move 0, m_ExtRect(Index).Y1, m_ButtonSize, m_ButtonSize
    End Select
    '-- Set tool tip text
    'On Error Resume Next
      lblTipRect.ToolTipText = m_uButton(Index).Tooltip 'm_Tooltip(Index - 1)
    'On Error GoTo 0
End Sub

'//

Private Sub pvEnableBar(ByVal bEnable As Boolean)

  Dim nBtn As Integer
    
    If (m_Count) Then
        '-- Enable/disable
        For nBtn = 1 To m_Count
            m_uButton(nBtn).Enabled = bEnable
        Next nBtn
        '-- Refresh
        Call pvRefresh
    End If
End Sub

Private Sub pvEnableButton(ByVal Index As Integer, ByVal Enabled As Boolean)
    
    If (m_Count) Then
        If (Index And Index <= m_Count And m_uButton(Index).Enabled <> Enabled) Then
            '-- Enable/disable
            With m_uButton(Index)
                .Enabled = Enabled
                 If (Not Enabled) Then .STATE = [btFlat]
            End With
            '-- Reset tooltip rect.
            lblTipRect.Move -lblTipRect.Width: m_LastOver = 0
            '-- Refresh
            Call pvRefresh(Index)
        End If
    End If
End Sub

'//

Private Sub pvRefresh(Optional ByVal Index As Integer = 0)

  Dim nBtn As Integer
    
    If (m_Count) Then
        If (Index = 0) Then
            '== All buttons...
            '-- Draw buttons
            For nBtn = 1 To m_Count
                Call pvPaintButton(nBtn)
                Call pvPaintBitmap(nBtn)
                If (IsRectEmpty(m_uButton(nBtn).Separator) = 0) Then
                    With m_uButton(nBtn).Separator
                        Select Case m_BarOrientation
                            Case [tbHorizontal]
                                Call pvDrawLine(.X1, .Y1, .X1, .Y2, m_hShadowPen)
                                Call pvDrawLine(.X2, .Y1, .X2, .Y2, m_hHighlightPen)
                            Case [tbVertical]
                                Call pvDrawLine(.X1, .Y1, .X2, .Y1, m_hShadowPen)
                                Call pvDrawLine(.X1, .Y2, .X2, .Y2, m_hHighlightPen)
                        End Select
                    End With
                End If
            Next nBtn
          Else
            '== Single button
            Call pvPaintButton(Index)
            Call pvPaintBitmap(Index)
        End If
        '-- Flat border [?]
        If (m_BarEdge) Then
            With m_BarRect
                Call pvDrawEdge(.X1, .Y1, .X2 - 1, .Y2 - 1, 0)
            End With
        End If
        
        '-- Refresh
        UserControl.Refresh
    End If
End Sub

Private Sub pvPaintButton(ByVal Index As Integer)
    
    '-- Background
    If (m_uButton(Index).Checked And m_uButton(Index).STATE = [btDown] And Not m_uButton(Index).Over) Then
        FillRect hdc, m_ClkRect(Index), m_hCheckBrush
      Else
        FillRect hdc, m_ExtRect(Index), m_hFaceBrush
    End If
    '-- Edge
    With m_ExtRect(Index)
        Select Case m_uButton(Index).STATE
            Case [btOver]
                Call pvDrawEdge(.X1, .Y1, .X2 - 1, .Y2 - 1, 0)
            Case [btDown]
                Call pvDrawEdge(.X1, .Y1, .X2 - 1, .Y2 - 1, -1)
        End Select
   End With
End Sub

Private Sub pvPaintBitmap(ByVal Index As Integer)
  
  Dim lOffset As Long
  
    '-- Image offset
    lOffset = 3 + (1 And m_uButton(Index).STATE = [btDown])
    '-- Paint masked bitmap
    With m_ExtRect(Index)
        Call pvDrawImage(Index, hdc, .X1 + lOffset, .Y1 + lOffset)
    End With
End Sub

Private Sub pvDrawImage(ByVal Index As Integer, ByVal hdc As Long, ByVal X As Integer, ByVal Y As Integer)

  Dim hIcon As Long

    If (m_uButton(Index).Enabled) Then
        '-- Normal
        ImageList_Draw m_hIL, Index - 1, hdc, X, Y, ILD_TRANSPARENT
      Else
        '-- Disabled
        hIcon = ImageList_GetIcon(m_hIL, Index - 1, 0)
        ImageList_Draw m_hIL, Index - 1, hdc, X, Y, ILD_TRANSPARENT
        DrawState hdc, m_hHighlightBrush, 0, hIcon, 0, X + 1, Y + 1, m_IconSize, m_IconSize, DST_ICON Or DSS_MONO 'DST_ICON Or DSS_DISABLED Or DSS_UNION Or DSS_MONO
        DrawState hdc, m_hShadowBrush, 0, hIcon, 0, X, Y, m_IconSize, m_IconSize, DST_ICON Or DSS_MONO 'DST_ICON Or DSS_DISABLED Or DSS_UNION Or DSS_MONO
        DestroyIcon hIcon
    End If
End Sub

Private Sub pvDrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal hPen As Long)

  Dim uPt     As POINTAPI
  Dim hOldPen As Long
            
    '-- Draw a simple line using given pen
    hOldPen = SelectObject(hdc, hPen)
    MoveToEx hdc, X1, Y1, uPt
    LineTo hdc, X2, Y2
    SelectObject hdc, hOldPen
End Sub

Private Sub pvDrawEdge(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal bPressed As Boolean)
                
    If (bPressed) Then
        '-- Button pressed
        Call pvDrawLine(X1, Y1, X2, Y1, m_hShadowPen)
        Call pvDrawLine(X1, Y1, X1, Y2, m_hShadowPen)
        Call pvDrawLine(X1, Y2, X2, Y2, m_hHighlightPen)
        Call pvDrawLine(X2, Y1, X2, Y2 + 1, m_hHighlightPen)
      Else
        '-- Mouse over button
        Call pvDrawLine(X1, Y1, X2, Y1, m_hHighlightPen)
        Call pvDrawLine(X1, Y1, X1, Y2, m_hHighlightPen)
        Call pvDrawLine(X1, Y2, X2, Y2, m_hShadowPen)
        Call pvDrawLine(X2, Y1, X2, Y2 + 1, m_hShadowPen)
    End If
End Sub

'//

Private Function pvTranslateColor(ByVal Clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    
    '-- OLE/RGB color to RGB color
    If (OleTranslateColor(Clr, hPal, pvTranslateColor)) Then
        pvTranslateColor = CLR_INVALID
    End If
End Function

Private Function pvShiftColor(ByVal Color As Long, ByVal Amount As Long) As Long

  Dim R As Long
  Dim b As Long
  Dim g As Long
    
    '-- Add amount
    R = (Color And &HFF) + Amount
    g = ((Color \ &H100) Mod &H100) + Amount
    b = ((Color \ &H10000) Mod &H100) + Amount
    '-- Check byte bounds
    If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
    If (g < 0) Then g = 0 Else If (g > 255) Then g = 255
    If (b < 0) Then b = 0 Else If (b > 255) Then b = 255
    
    '-- Return shifted color
    pvShiftColor = R + 256& * g + 65536 * b
End Function

Private Sub pvGetColors(ByVal FaceColor As OLE_COLOR)
    
  Dim lFaceColor As Long
    
    '-- Get long value
    lFaceColor = pvTranslateColor(FaceColor)
    
    '-- Build brush and pens
    If (m_hFaceBrush) Then
        DeleteObject m_hFaceBrush
        m_hFaceBrush = 0
    End If
    If (m_hHighlightPen) Then
        DeleteObject m_hHighlightPen
        m_hHighlightPen = 0
    End If
    If (m_hShadowPen) Then
        DeleteObject m_hShadowPen
        m_hShadowPen = 0
    End If
    If (m_hHighlightBrush) Then
        DeleteObject m_hHighlightBrush
        m_hHighlightBrush = 0
    End If
    If (m_hShadowBrush) Then
        DeleteObject m_hShadowBrush
        m_hShadowBrush = 0
    End If
    
    m_hFaceBrush = CreateSolidBrush(lFaceColor)
    m_hHighlightPen = CreatePen(PS_SOLID, 1, pvShiftColor(lFaceColor, &H2F))
    m_hShadowPen = CreatePen(PS_SOLID, 1, pvShiftColor(lFaceColor, -&H40))
    m_hHighlightBrush = CreateSolidBrush(pvShiftColor(lFaceColor, &H2F))
    m_hShadowBrush = CreateSolidBrush(pvShiftColor(lFaceColor, -&H40))
    
    '-- For check effect
    UserControl.ForeColor = pvShiftColor(lFaceColor, &H2F)
End Sub

'//

Private Sub pvUpdateButtonState(ByVal Index As Integer, _
                                ByVal InButton As Boolean, _
                                ByVal MouseButton As MouseButtonConstants, _
                                ByVal MouseEvent As tbMouseEventConstants)
  
  Dim uTmpButton As tButton
    
    '-- Store current button state / Over button [?]
    uTmpButton = m_uButton(Index)
    m_uButton(Index).Over = InButton
    
    '-- Check new state
    With m_uButton(Index)
        
        Select Case MouseEvent
            
            Case [btMouseDown] '-- Mouse pressed
                If (MouseButton = vbLeftButton) Then
                    .STATE = [btDown]
                End If
                
            Case [btMouseMove] '-- Mouse moving
                If (InButton) Then
                    If (MouseButton = vbLeftButton) Then
                        .STATE = [btDown]
                      Else
                        If (Not .Checked) Then
                            .STATE = [btOver]
                        End If
                        tmrTip.Enabled = -1
                    End If
                  Else
                    If (Not .Checked) Then
                        .STATE = [btFlat]
                    End If
                End If
                
            Case [btMouseUp]  '-- Mouse released
                If (InButton) Then
                    If (MouseButton = vbLeftButton) Then
                        Select Case .Type
                            Case [btNormal]
                                .STATE = [btOver]
                            Case [btCheck]
                                .Checked = Not .Checked
                                .STATE = -.Checked * [btDown]
                            Case [btOption]
                                .Checked = -1
                                .STATE = [btDown]
                                 Call pvUpdateOptionButtons(Index)
                        End Select
                      Else
                        If (Not .Checked And MouseButton = vbEmpty) Then
                            .STATE = [btFlat]
                        End If
                    End If
                End If
        End Select
        
        '-- Refresh [?]
        If (.STATE <> uTmpButton.STATE Or .Checked <> uTmpButton.Checked Or .Over <> uTmpButton.Over) Then
            Call pvRefresh(Index)
        End If
        '-- Raise [Click]/[Check] event [?]
        If (InButton And MouseEvent = [btMouseUp]) Then
            Select Case m_uButton(Index).Type
                Case [btNormal]
                    RaiseEvent ButtonClick(Index, MouseButton, m_ExtRect(Index).X1, m_ExtRect(Index).Y1)
                Case [btCheck], [btOption]
                    RaiseEvent ButtonClick(Index, MouseButton, m_ExtRect(Index).X1, m_ExtRect(Index).Y1)
                    If (.Checked <> uTmpButton.Checked) Then
                        RaiseEvent ButtonCheck(Index, m_ExtRect(Index).X1, m_ExtRect(Index).Y1)
                    End If
            End Select
        End If
    End With
End Sub

Private Sub pvUpdateOptionButtons(ByVal CurrentIndex As Integer)

  Dim nIdx As Integer
    
    '-- Right/below buttons
    nIdx = CurrentIndex
    Do While nIdx < m_Count
        If (IsRectEmpty(m_uButton(nIdx).Separator) = 0) Then
            Exit Do
          Else
            nIdx = nIdx + 1
            With m_uButton(nIdx)
                If (.Type = [btOption] And .Checked) Then
                    .Checked = 0
                    .STATE = [btFlat]
                     Call pvRefresh(nIdx)
                End If
            End With
        End If
    Loop
    
    '-- Left/above buttons
    nIdx = CurrentIndex
    Do While nIdx > 1
        nIdx = nIdx - 1
        If (IsRectEmpty(m_uButton(nIdx).Separator) = 0) Then
            Exit Do
          Else
            With m_uButton(nIdx)
                If (.Type = [btOption] And .Checked) Then
                    .Checked = 0
                    .STATE = [btFlat]
                     Call pvRefresh(nIdx)
                End If
            End With
        End If
    Loop
End Sub

'//

Private Sub tmrTip_Timer()
  
  Dim uPt As POINTAPI
    
    '-- Cursor out of toolbar [?]
    GetCursorPos uPt
    If (WindowFromPoint(uPt.X, uPt.Y) <> hWnd) Then
        '-- Disable timer and refresh
        tmrTip.Enabled = 0
        If m_LastOver > 0 Then
        Call pvUpdateButtonState(m_LastOver, 0, 0, [btMouseMove])
        End If
    End If
End Sub

'==================================================================================================
' Properties
'==================================================================================================

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call pvGetColors(New_BackColor)
    Call pvRefresh
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call pvEnableBar(New_Enabled)
End Property

Public Property Get BarOrientation() As tbOrientationConstants
    BarOrientation = m_BarOrientation
End Property
Public Property Let BarOrientation(ByVal New_BarOrientation As tbOrientationConstants)
   ' If (Not Ambient.UserMode) Then
        m_BarOrientation = New_BarOrientation
   ' End If
End Property

Public Property Get BarEdge() As Boolean
    BarEdge = m_BarEdge
End Property
Public Property Let BarEdge(ByVal New_BarEdge As Boolean)
    m_BarEdge = New_BarEdge
    Call pvRefresh
End Property

Public Property Get ButtonsCount() As Integer
    ButtonsCount = m_Count
End Property

'//

Private Sub UserControl_InitProperties()
    m_BarOrientation = m_def_BarOrientation
    m_BarEdge = m_def_BarEdge
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        UserControl.BackColor = .ReadProperty("BackColor", &H8000000F)
        UserControl.Enabled = .ReadProperty("Enabled", -1)
        m_BarOrientation = .ReadProperty("BarOrientation", m_def_BarOrientation)
        m_BarEdge = .ReadProperty("BarEdge", m_def_BarEdge)
    End With
    Call pvGetColors(UserControl.BackColor)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
        Call .WriteProperty("Enabled", UserControl.Enabled, -1)
        Call .WriteProperty("BarOrientation", m_BarOrientation, m_def_BarOrientation)
        Call .WriteProperty("BarEdge", m_BarEdge, m_def_BarEdge)
    End With
End Sub

Public Function GetToolTips(Index As Long) As String
       GetToolTips = m_uButton(Index).Tooltip
End Function
