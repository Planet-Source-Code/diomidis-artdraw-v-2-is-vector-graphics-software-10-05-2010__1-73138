Attribute VB_Name = "ModApi"
'(c) 2007 Diomidisk
Option Explicit

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type POINTAPIs
    X As Single
    Y As Single
End Type
Public Declare Function PolyBezier Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

'Region
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (ByRef lpRect As RECT) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectClipPath Lib "gdi32.dll" (ByVal hdc As Long, ByVal iMode As Long) As Long
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetRegionData Lib "gdi32.dll" (ByVal hRgn As Long, ByVal dwCount As Long, ByRef lpRgnData As Any) As Long
Public Declare Function ExtCreateRegion Lib "gdi32" (ByRef lpXform As Any, ByVal nCount As Long, ByRef lpRgnData As Any) As Long
Public Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Public Declare Function ModifyWorldTransform Lib "gdi32" (ByVal hdc As Long, lpXform As XForm, ByVal iMode As Long) As Long

'Transform
Public Declare Function SetWorldTransform Lib "gdi32" (ByVal hdc As Long, ByRef lpXform As XForm) As Long
Public Declare Function GetWorldTransform Lib "gdi32" (ByVal hdc As Long, ByRef lpXform As XForm) As Long
Public Declare Function CombineTransform Lib "gdi32" (ByRef lpXFormResult As XForm, ByRef lpXForm1 As XForm, ByRef lpXForm2 As XForm) As Long
Public Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, ByRef lpPoint As Any) As Long
Public Declare Function SetViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, ByRef lpSize As Any) As Long
Public Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, ByRef lpPoint As Any) As Long
Public Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, ByRef lpSize As Any) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As MMETRIC) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function InflateRect Lib "user32" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetGraphicsMode Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Public Declare Function GetViewportExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As POINTAPI) As Long
Public Declare Function OffsetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, ByRef lpPoint As Any) As Long
Public Declare Function GetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, lpSize As POINTAPI) As Long
Public Declare Function GetObjectApi Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function DPtoLP Lib "gdi32" (ByVal hdc As Long, ByRef lpPoint As Any, ByVal nCount As Long) As Long
Public Declare Function LPtoDP Lib "gdi32" (ByVal hdc As Long, ByRef lpPoint As Any, ByVal nCount As Long) As Long

Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public Const RGN_DIFF = 4

Public Enum FillMode
       fALTERNATE = 1 ' ALTERNATE and WINDING are
       fWINDING = 2 ' constants for FillMode.
End Enum

'Path
Public Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function FillPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CloseFigure Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long

Public Type SIZEL
    cx As Long
    cy As Long
End Type

Public Type typScaleMode
    smViewportOrigin As POINTAPI
    smViewportExtent As SIZEL
    smWindowOrigin As POINTAPI
    smWindowExtent As SIZEL
    smMappingMode As Long
End Type

Public Const TA_TOP As Long = &H0
Public Const TA_LEFT As Long = &H0
Public Const TA_RIGHT As Long = &H2
Public Const TA_BOTTOM As Long = &H8

Public Type XForm
    eM11 As Single
    eM12 As Single
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type

'This Enum is needed to set the "Mapping" property for EMF images
Public Enum MMETRIC
    MM_TEXT = 1
    MM_LOMETRIC = 2
    MM_HIMETRIC = 3
    MM_LOENGLISH = 4
    MM_HIENGLISH = 5
    MM_TWIPS = 6
    MM_ISOTROPIC = 7
    MM_ANISOTROPIC = 8
    MM_ADLIB = 9
End Enum

Public Const MWT_IDENTITY = 1
Public Const MWT_LEFTMULTIPLY = 2
Public Const MWT_RIGHTMULTIPLY = 3

Public Const GM_ADVANCED = &H2

Public Declare Function GetTickCount Lib "kernel32" () As Long
'Public Declare Function SetRect Lib "User32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dX As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
'Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Public Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Public Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function CreateScalableFontResource Lib "gdi32" Alias "CreateScalableFontResourceA" (ByVal fHidden As Long, ByVal lpszResourceFile As String, ByVal lpszFontFile As String, ByVal lpszCurrentPath As String) As Long

Public Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal imageType As Long, ByVal NewWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetBrushOrgEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, ByRef lpPt As Any) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LogBrush) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Public Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, ByRef lplb As LogBrush, ByVal dwStyleCount As Long, ByRef lpStyle As Long) As Long

Public Type LogBrush ' 12 bytes
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

' Pen styles
Public Const PS_SOLID As Long = 0
Public Const PS_DASH As Long = 1
Public Const PS_DOT As Long = 2
Public Const PS_NULL As Long = 5
Public Const PS_INSIDEFRAME As Long = 6
Public Const PS_USERSTYLE As Long = 7
Public Const PS_ALTERNATE As Long = 8
Public Const PS_STYLE_MASK As Long = &HF

Public Const PS_ENDCAP_ROUND As Long = &H0
Public Const PS_ENDCAP_SQUARE As Long = &H100
Public Const PS_ENDCAP_FLAT As Long = &H200
Public Const PS_ENDCAP_MASK As Long = &HF00

Public Const PS_JOIN_ROUND As Long = &H0
Public Const PS_JOIN_BEVEL As Long = &H1000
Public Const PS_JOIN_MITER As Long = &H2000
Public Const PS_JOIN_MASK As Long = &HF000&

Public Const PS_COSMETIC As Long = &H0
Public Const PS_GEOMETRIC As Long = &H10000

'Fill Style
Public Const BS_SOLID = 0
Public Const BS_NULL = 1
Public Const BS_HOLLOW = BS_NULL
Public Const BS_HATCHED = 2
Public Const HS_BDIAGONAL = 3
Public Const HS_CROSS = 4
Public Const HS_DIAGCROSS = 5
Public Const HS_FDIAGONAL = 2
Public Const HS_HORIZONTAL = 0
Public Const HS_VERTICAL = 1

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function PolyDraw Lib "gdi32" (ByVal hdc As Long, ByRef lpPt As POINTAPI, ByRef lpbTypes As Byte, ByVal cCount As Long) As Long
Public Declare Function PolyBezierTo Lib "gdi32.dll" (ByVal hdc As Long, ByRef lpPt As POINTAPI, ByVal cCount As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Public Declare Function RectAngleEx Lib "gdi32.dll" Alias "Rectangle" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
 
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function TransparentBlt Lib "MSImg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateDCA Lib "gdi32" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, ByRef lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Public Declare Function CloseMetaFile Lib "gdi32" (ByVal hMf As Long) As Long
Public Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMf As Long) As Long
'Public Declare Function SetWindowExtEx Lib "GDI32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, ByRef lpSize As Any) As Long
Public Declare Function Escape Lib "gdi32" (ByVal hdc As Long, ByVal nEscape As Long, ByVal nCount As Long, ByVal lpInData As String, lpOutData As Any) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dy As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByRef lpBits As Any, ByRef BInfo As BitmapInfoHeader, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByRef pBitmapInfo As BitmapInfoHeader, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Public Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
' for popup menu
Public Const MF_CHECKED = &H8&
Public Const MF_APPEND = &H100&
Public Const TPM_LEFTALIGN = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_GRAYED = &H1&
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const TPM_RETURNCMD = &H100&
Public Const TPM_RIGHTBUTTON = &H2&

Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal hwnd As Long, ByVal lptpm As Any) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public hMenu As Long

Type Bitmap
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Type BitmapInfoHeader   '40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Type BITMAPINFO
  bmiHeader As BitmapInfoHeader
  bmiColors As RGBQUAD
End Type
