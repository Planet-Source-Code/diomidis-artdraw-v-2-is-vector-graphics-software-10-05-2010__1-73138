Attribute VB_Name = "ModHookMouse"
 Option Explicit
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type MSG
  hwnd As Long
  message As Long
  wParam As Long
  lParam As Long
  time As Long
  pt As POINTAPI
End Type
Public Declare Function CallNextHookEx& Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Integer, lParam As Any)
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Public Declare Function SetWindowsHookEx& Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long)
Public Declare Function UnhookWindowsHookEx& Lib "user32" (ByVal hHook As Long)

Public Const WH_GETMESSAGE = 3
Public Const MSH_MOUSEWHEEL = "MSWHEEL_ROLLMSG"
Public IMWHEEL_MSG As Long
Public HWND_HOOK As Long
Private Fform As Form

Public Function IMWheel(ByVal nCode As Long, ByVal wParam As Long, lParam As MSG) As Long
   Debug.Print lParam.message
  If lParam.message = IMWHEEL_MSG Then
      Fform.WheelMoved lParam.wParam, lParam.pt.X, lParam.pt.Y
  End If
  IMWheel = CallNextHookEx(HWND_HOOK, nCode, wParam, lParam)
End Function

Public Function IMWheel_Hook(F_form As Form) As Long
     Set Fform = F_form
     IMWHEEL_MSG = RegisterWindowMessage(MSH_MOUSEWHEEL)
     HWND_HOOK = SetWindowsHookEx(WH_GETMESSAGE, AddressOf IMWheel, 0, GetCurrentThreadId)
End Function

Public Sub IMWheel_Unhook()
     UnhookWindowsHookEx HWND_HOOK
End Sub
