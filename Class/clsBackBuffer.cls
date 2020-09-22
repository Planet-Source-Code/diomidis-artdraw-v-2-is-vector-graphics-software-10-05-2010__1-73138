VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBackBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As BitmapInfoHeader, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDesktopWindow Lib "User32.dll" () As Long
Private Declare Function GetDC Lib "User32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "User32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function AlphaBlend Lib "MSImg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal blendFunction As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
'Private Declare Function GetObjectApi Lib "GDI32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Sub RtlMoveMemory Lib "Kernel32.dll" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Type BitmapInfoHeader ' 40 bytes
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

Private Type Bitmap ' 24 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Dim m_hDC As Long
Dim m_hDIB As Long, hOldDIB As Long
Dim m_Width As Long
Dim m_Height As Long
Dim m_UseInternalBlend As Boolean

Dim ScanLine As Long
Dim DataPtr As Long
Dim BitDepth As Byte

Private Const SHADEBLENDCAPS As Long = 120 ' Shading and blending caps
Private Const BITSPIXEL As Long = &HC      ' Number of bits per pixel
Private Const SB_CONST_ALPHA As Long = &H1
Private Const BI_RGB As Long = &H0
Private Const HALFTONE As Long = &H4
Private Const OBJ_BITMAP As Long = &H7

Public Property Get hDC() As Long
    hDC = m_hDC
End Property

Public Property Get Width() As Long
    Width = m_Width
End Property
Public Property Let Width(ByVal inNew As Long)
    If ((inNew < 0) Or (inNew = m_Width)) Then Exit Property
    Call ReSize(inNew, m_Height)
End Property

Public Property Get Height() As Long
    Height = m_Height
End Property
Public Property Let Height(ByVal inNew As Long)
    If ((inNew < 0) Or (inNew = m_Height)) Then Exit Property
    Call ReSize(m_Width, inNew)
End Property

Public Property Get UseInternalBlend() As Boolean
    UseInternalBlend = m_UseInternalBlend
End Property
Public Property Let UseInternalBlend(ByVal inNew As Boolean)
    Dim DeskDC As Long
    
    DeskDC = GetDC(0) ' Check to see if this device supports source-constant alpha blending
    m_UseInternalBlend = inNew Or (Not CBool(GetDeviceCaps(DeskDC, SHADEBLENDCAPS) And SB_CONST_ALPHA))
    Call ReleaseDC(0, DeskDC)
End Property

Public Function Create(ByVal inWidth As Long, ByVal inHeight As Long) As Boolean
    Dim DeskWnd As Long, DeskDC As Long, DeskDepth As Long
    Dim DIBInf As BitmapInfoHeader
    
    ' Destroy any current back-buffer and associated GDI objects
    Call Destroy
    
    ' Validate buffer size
    If ((inWidth < 1) Or (inHeight < 1)) Then Exit Function
    
    ' Get desktop bit-display
    DeskWnd = GetDesktopWindow()
    DeskDC = GetDC(DeskWnd)
    DeskDepth = GetDeviceCaps(DeskDC, BITSPIXEL)
    Call ReleaseDC(DeskWnd, DeskDC)
    
    ' Create DIB information structure
    With DIBInf
        .biSize = Len(DIBInf)
        .biWidth = inWidth
        .biHeight = inHeight
        .biPlanes = 1 ' Match desktop at 32 or 24 bit
        .biBitCount = IIf(DeskDepth = 32, 32, 24)
        .biCompression = BI_RGB
        
        ' Calculate DWord aligned scanline width
        ScanLine = (((.biWidth * (.biBitCount \ 8)) + 3) \ 4) * 4
        .biSizeImage = ScanLine * .biHeight
    End With
    
    ' Create device context
    m_hDC = CreateCompatibleDC(0)
    
    If (m_hDC) Then ' Create DIBSection
        m_hDIB = CreateDIBSection(m_hDC, DIBInf, 0, DataPtr, 0, 0)
        
        If (m_hDIB) Then ' Select DIBSection into DC
            hOldDIB = SelectObject(m_hDC, m_hDIB)
            Create = CBool(hOldDIB) ' Validate selection
        End If
    End If
    
    If (Not Create) Then
        ' Something went wrong, tear down
        Call Destroy
    Else ' Expose properties
        m_Width = inWidth
        m_Height = inHeight
        BitDepth = DIBInf.biBitCount
    End If
End Function

Public Function CreateSnapshop(ByVal inDC As Long, ByVal inX As Long, ByVal inY As Long, ByVal inWidth As Long, ByVal inHeight As Long) As Boolean
    ' Create a new back-buffer and fill it with a snapshot from the DC
    If (Me.Create(inWidth, inHeight)) Then CreateSnapshop = Me.SnapShot(inDC, inX, inY)
End Function

Public Function ReSize(ByVal inWidth As Long, ByVal inHeight As Long, Optional ByVal inStretch As Boolean = False) As Boolean
    Dim TempCopy As clsBackBuffer
    Dim OldStretch As Long
    
    ' Create a new back-buffer
    Set TempCopy = New clsBackBuffer
    
    If (TempCopy.Create(inWidth, inHeight)) Then
        If (inStretch) Then ' Stretch current to new size
            OldStretch = SetStretchBltMode(TempCopy.hDC, HALFTONE)
            ReSize = StretchBlt(TempCopy.hDC, 0, 0, inWidth, inHeight, m_hDC, 0, 0, m_Width, m_Height, vbSrcCopy) <> 0
            Call SetStretchBltMode(TempCopy.hDC, OldStretch)
        Else ' Crop current to new size
            ReSize = BitBlt(TempCopy.hDC, 0, 0, Min(m_Width, inWidth), Min(m_Height, inHeight), m_hDC, 0, 0, vbSrcCopy) <> 0
        End If
        
        If (ReSize) Then ' Re-scale local buffer and copy new image back
            If (Me.Create(inWidth, inHeight)) Then _
                ReSize = BitBlt(m_hDC, 0, 0, inWidth, inHeight, TempCopy.hDC, 0, 0, vbSrcCopy) <> 0
        End If
    End If
    
    ' Clean up temp back-buffer
    Set TempCopy = Nothing
End Function

Public Function Clone() As clsBackBuffer
    Set Clone = New clsBackBuffer
    
    ' Create new back buffer the same size as this one
    If (Clone.Create(m_Width, m_Height)) Then ' Copy current image to clone
        If (BitBlt(Clone.hDC, 0, 0, m_Width, m_Height, m_hDC, 0, 0, vbSrcCopy) = 0) Then _
            Set Clone = Nothing
    Else
        Set Clone = Nothing
    End If
End Function

Public Function SnapShot(ByVal inDC As Long, ByVal inX As Long, ByVal inY As Long) As Boolean
    If (m_hDC = 0) Then Exit Function ' Grab a snapshot from this DC
    SnapShot = BitBlt(m_hDC, 0, 0, m_Width, m_Height, inDC, inX, inY, vbSrcCopy) <> 0
End Function

Public Function Draw(ByVal inDC As Long, _
                    Optional ByVal inX As Long = 0, _
                    Optional ByVal inY As Long = 0, _
                    Optional ByVal inBlend As Byte = &HFF) As Boolean
    
    If (inBlend = &HFF) Then ' Full Blend, just draw
        Draw = BitBlt(inDC, inX, inY, m_Width, m_Height, m_hDC, 0, 0, vbSrcCopy) <> 0
    Else ' Semi-Blend
        If (m_UseInternalBlend) Then ' Use internal blending
            Dim BlendBuffer As clsBackBuffer
            
            ' Create blended buffer
            Set BlendBuffer = InternalBlend(inDC, inX, inY, inBlend)
            If (Not (BlendBuffer Is Nothing)) Then Draw = BlendBuffer.Draw(inDC, inX, inY, &HFF) ' Draw to target DC
            Set BlendBuffer = Nothing ' Clean up back buffer
        Else ' Use system blending
             Dim lngBlend As Long
             lngBlend = Val("&h" & Hex(inBlend) & "00" & "00")
             Draw = AlphaBlend(inDC, inX, inY, m_Width, m_Height, m_hDC, 0, 0, m_Width, m_Height, lngBlend)
            'Draw = AlphaBlend(inDC, inX, inY, m_Width, m_Height, m_hDC, 0, 0, m_Width, m_Height, inBlend * &H10000) <> 0
        End If
    End If
End Function

Private Function InternalBlend(ByVal inDC As Long, ByVal inX As Long, ByVal inY As Long, ByVal inBlend As Byte) As clsBackBuffer
    Dim Background As clsBackBuffer
    Dim hBackDIB As Long
    Dim BackInf As Bitmap
    Dim LocalData() As Byte, BackData() As Byte
    Dim LocalPos As Long, LocalStep As Long
    Dim BackPos As Long, BackStep As Long
    Dim LoopX As Long, LoopY As Long
    Dim BlendAmt As Single
    
    ' Check internal flag
    If (Not m_UseInternalBlend) Then Exit Function
    
    If (inBlend = &HFF) Then ' Just clone
        Set InternalBlend = Me.Clone()
        Exit Function
    End If
    
    ' Create new back-buffer
    Set Background = New clsBackBuffer
    
    ' Create a snapshot of the current target DC
    If (Background.CreateSnapshop(inDC, inX, inY, m_Width, m_Height)) Then
        If (inBlend) Then ' Anything other than 0%...
            hBackDIB = GetCurrentObject(Background.hDC, OBJ_BITMAP) ' Get back-buffer's Bitmap
            If (GetObjectApi(hBackDIB, Len(BackInf), BackInf)) Then    ' Get information about Bitmap
                ' Allocate and read image data buffers locally
                ReDim BackData((BackInf.bmWidthBytes * BackInf.bmHeight) - 1) As Byte
                ReDim LocalData((ScanLine * m_Height) - 1) As Byte
                Call RtlMoveMemory(BackData(0), ByVal BackInf.bmBits, UBound(BackData()) + 1)
                Call RtlMoveMemory(LocalData(0), ByVal DataPtr, UBound(LocalData()) + 1)
                
                ' Calculate byte per pixel counts for both buffers
                BackStep = BackInf.bmBitsPixel \ 8
                LocalStep = BitDepth \ 8
                BlendAmt = inBlend / &HFF ' Floating point Blend
                
                ' Blend loop
                For LoopY = 0 To m_Height - 1
                    LocalPos = LoopY * ScanLine
                    BackPos = LoopY * BackInf.bmWidthBytes
                    
                    For LoopX = 0 To m_Width - 1 ' Blend source and destination pixels using linear interpolation
                        BackData(BackPos) = (BackData(BackPos) * (1 - BlendAmt)) + (LocalData(LocalPos) * BlendAmt)
                        BackData(BackPos + 1) = (BackData(BackPos + 1) * (1 - BlendAmt)) + (LocalData(LocalPos + 1) * BlendAmt)
                        BackData(BackPos + 2) = (BackData(BackPos + 2) * (1 - BlendAmt)) + (LocalData(LocalPos + 2) * BlendAmt)
                        
                        ' Increment data position
                        LocalPos = LocalPos + LocalStep
                        BackPos = BackPos + BackStep
                    Next LoopX
                Next LoopY
                
                ' Push blended data back into background DIB
                Call RtlMoveMemory(ByVal BackInf.bmBits, BackData(0), UBound(BackData()) + 1)
            Else ' Couldn't get Bitmap information
                Set Background = Nothing
            End If
        End If
    Else ' Failed to create snapshot buffer
        Set Background = Nothing
    End If
    
    ' Return blended buffer
    Set InternalBlend = Background
End Function

Private Sub Destroy() ' Clean up GDI objects
    If (hOldDIB) Then Call SelectObject(m_hDC, m_hDIB)
    If (m_hDC) Then Call DeleteDC(m_hDC)
    If (m_hDIB) Then Call DeleteObject(m_hDIB)
    
    ' Re-set member variables
    m_hDC = 0
    m_hDIB = 0
    m_Width = 0
    m_Height = 0
    BitDepth = 0
    ScanLine = 0
    hOldDIB = 0
End Sub

Private Function Min(ByVal inA As Long, ByVal inB As Long) As Long
    If (inA < inB) Then Min = inA Else Min = inB ' Return the minimum of two values
End Function

Private Sub Class_Initialize() ' By default try to use (hadrware accelerated) system blending
    Me.UseInternalBlend = False
End Sub

Private Sub Class_Terminate()
    Call Destroy ' Automatically clean up before quitting
End Sub
