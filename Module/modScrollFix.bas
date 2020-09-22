Attribute VB_Name = "modScrollFix"
'**********************************************************************************************************************'
' Module    : modScrollFix
' Author    : Joseph M. Ferris <josephmferris@cox.net>
' Date      : 05.12.2003
' Depends   : None.
' Purpose   : Provides a fix for the VB Scrollbar controls on Windows NT-derived platforms by subclassing to listen
'             for the WM_CTLCOLORSCROLLBAR message which erronously draws a white background on the scrollbar.
'
' Notes     : 1.  Based upon Knowledge Base sample
'             2.  Applies to all scrollbars for a given hWnd.
'**********************************************************************************************************************'

Option Explicit

'**********************************************************************************************************************'
'
' API Declarations - USER32.DLL
'
'**********************************************************************************************************************'

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'**********************************************************************************************************************'
'
' Constant Declarations
'
'**********************************************************************************************************************'

Private Const GWL_WNDPROC                As Long = -4
Private Const WM_CTLCOLORSCROLLBAR       As Long = 311

'**********************************************************************************************************************'
'
' Private Member Declarations
'
'**********************************************************************************************************************'

Private m_lngPrevHwnd                   As Long

'**********************************************************************************************************************'
'
' Public Member Declarations
'
'**********************************************************************************************************************'

Public g_lngTargetHwnd                  As Long

'**********************************************************************************************************************'
' Procedure : Hook
' Date      : 05.12.2003
' Purpose   : Subclasses the currently specified handle to watch for the WM_CTLCOLORSCROLLBAR message.
' Input     : None.
' Output    : None.
'**********************************************************************************************************************'
Public Sub Hook()

    ' Make sure that there is a handle to hook into.
    If Not (g_lngTargetHwnd = 0) Then
    
        ' Set the location to relay Windows messages to.  The return will be the old address.
        m_lngPrevHwnd = SetWindowLong(g_lngTargetHwnd, GWL_WNDPROC, AddressOf WindowProc)

    End If
    
End Sub

'**********************************************************************************************************************'
' Procedure : Unhook
' Date      : 05.12.2003
' Purpose   : Terminates subclassing for the currently specified handle and returns it to the orignal processes.
' Input     : None.
' Output    : None.
'**********************************************************************************************************************'
Public Sub Unhook()
   
Dim lngResult As Long
   
    ' Make sure that there is at least one hWnd between the new and old, to ensure that there actually is something
    ' to unhook.
    If m_lngPrevHwnd = 0 Or g_lngTargetHwnd = 0 Then
        Exit Sub
    End If
    
    ' Resume message handling to the initial address.
    '
    lngResult = SetWindowLong(g_lngTargetHwnd, GWL_WNDPROC, m_lngPrevHwnd)
   
End Sub

'**********************************************************************************************************************'
' Procedure : WindowProc
' Date      : 05.12.2003
' Purpose   :
' Input     : hw          - Destination handle for the message.
'             uMsg        - Message.
'             wParam      - Message parameter first parameter set.
'             lParam      - Message parameter second parameter set.
' Output    : WindowProc  - Result of calling the original handler.
'**********************************************************************************************************************'
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
                    ByVal wParam As Long, ByVal lParam As Long) As Long

    ' Look for the WM_CTLCOLORSCROLLBAR message.  This message is only present on NT-derived system and draws the
    ' scrollbar in an awful white background.  Just ignore it.
    If Not (uMsg = WM_CTLCOLORSCROLLBAR) Then
        
        ' Pass the message on to the original handler.
        WindowProc = CallWindowProc(m_lngPrevHwnd, hw, uMsg, wParam, lParam)
        
    End If

End Function




