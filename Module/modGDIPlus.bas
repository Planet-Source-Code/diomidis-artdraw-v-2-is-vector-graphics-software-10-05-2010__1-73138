Attribute VB_Name = "modGDIPlus"
Option Explicit

Private Type GUID
   Data1    As Long
   Data2    As Integer
   Data3    As Integer
   Data4(7) As Byte
End Type

Private Type PICTDESC
   Size     As Long
   Type     As Long
   hBmp     As Long
   hPal     As Long
   Reserved As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type PWMFRect16
    Left   As Integer
    Top    As Integer
    Right  As Integer
    Bottom As Integer
End Type

Private Type wmfPlaceableFileHeader
    Key         As Long
    hMf         As Integer
    BoundingBox As PWMFRect16
    Inch        As Integer
    Reserved    As Long
    CheckSum    As Integer
End Type

' GDI Functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

' GDI+ functions
Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal Filename As Long, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hdc As Long, GpGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hBmp As Long, ByVal hPal As Long, GpBitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipCreateMetafileFromWmf Lib "gdiplus.dll" (ByVal hWmf As Long, ByVal deleteWmf As Long, WmfHeader As wmfPlaceableFileHeader, Metafile As Long) As Long
Private Declare Function GdipCreateMetafileFromEmf Lib "gdiplus.dll" (ByVal hEmf As Long, ByVal deleteEmf As Long, Metafile As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus.dll" (ByVal hIcon As Long, GpBitmap As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal Token As Long)

' GDI and GDI+ constants
Private Const PLANES = 14            '  Number of planes
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const PATCOPY = &HF00021     ' (DWORD) dest = pattern
Private Const PICTYPE_BITMAP = 1     ' Bitmap type
Private Const InterpolationModeHighQualityBicubic = 7
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Private Const UnitPixel = 2
Private Token As Long
 
' Initialises GDI Plus
Public Function InitGDIPlus() As Long
    Dim Token    As Long
    Dim gdipInit As GdiplusStartupInput
    
    gdipInit.GdiplusVersion = 1
    GdiplusStartup Token, gdipInit, ByVal 0&
    InitGDIPlus = Token
End Function

' Frees GDI Plus
Public Sub FreeGDIPlus(Token As Long)
    GdiplusShutdown Token
End Sub

' Loads the picture (optionally resized)
Public Function LoadPictureGDIPlus(PicFile As String, _
                                   Optional Width As Long = -1, _
                                   Optional Height As Long = -1, _
                                   Optional ByVal BackColor As Long = vbWhite, _
                                   Optional RetainRatio As Boolean = False) As IPicture
    Dim hdc     As Long
    Dim hBitmap As Long
    Dim Img     As Long
    
    ' Initialise GDI+
    Token = InitGDIPlus
   If Token <> 0 Then
    ' Load the image
    If GdipLoadImageFromFile(StrPtr(PicFile), Img) <> 0 Then
        Err.Raise 999, "GDI+ Module", "Error loading picture " & PicFile
        Exit Function
    End If
    
    ' Calculate picture's width and height if not specified
    If Width = -1 Or Height = -1 Then
        GdipGetImageWidth Img, Width
        GdipGetImageHeight Img, Height
    End If
    
    ' Initialise the hDC
    InitDC hdc, hBitmap, BackColor, Width, Height

    ' Resize the picture
    gdipResize Img, hdc, Width, Height, RetainRatio
    GdipDisposeImage Img
    
    ' Get the bitmap back
    GetBitmap hdc, hBitmap

    ' Create the picture
    Set LoadPictureGDIPlus = CreatePicture(hBitmap)
    
    ' Free GDI+
    FreeGDIPlus Token
  Else
          MsgBox " Error " + Str(Token) + " : " + GdiErrorString(Token), vbCritical
  End If
End Function

' Initialises the hDC to draw
Private Sub InitDC(hdc As Long, hBitmap As Long, BackColor As Long, Width As Long, Height As Long)
    Dim hBrush As Long
        
    ' Create a memory DC and select a bitmap into it, fill it in with the backcolor
    hdc = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hdc, PLANES), GetDeviceCaps(hdc, BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(hdc, hBitmap)
    hBrush = CreateSolidBrush(BackColor)
    hBrush = SelectObject(hdc, hBrush)
    PatBlt hdc, 0, 0, Width, Height, PATCOPY
    DeleteObject SelectObject(hdc, hBrush)
End Sub

' Resize the picture using GDI plus
Private Sub gdipResize(Img As Long, hdc As Long, Width As Long, Height As Long, Optional RetainRatio As Boolean = False)
    Dim Graphics   As Long      ' Graphics Object Pointer
    Dim OrWidth    As Long      ' Original Image Width
    Dim OrHeight   As Long      ' Original Image Height
    Dim OrRatio    As Double    ' Original Image Ratio
    Dim DesRatio   As Double    ' Destination rect Ratio
    Dim DestX      As Long      ' Destination image X
    Dim DestY      As Long      ' Destination image Y
    Dim DestWidth  As Long      ' Destination image Width
    Dim DestHeight As Long      ' Destination image Height
    
    GdipCreateFromHDC hdc, Graphics
    GdipSetInterpolationMode Graphics, InterpolationModeHighQualityBicubic
    
    If RetainRatio Then
        GdipGetImageWidth Img, OrWidth
        GdipGetImageHeight Img, OrHeight
        
        OrRatio = OrWidth / OrHeight
        DesRatio = Width / Height
        
        ' Calculate destination coordinates
        DestWidth = IIf(DesRatio < OrRatio, Width, Height * OrRatio)
        DestHeight = IIf(DesRatio < OrRatio, Width / OrRatio, Height)
        DestX = (Width - DestWidth) / 2
        DestY = (Height - DestHeight) / 2

        GdipDrawImageRectRectI Graphics, Img, DestX, DestY, DestWidth, DestHeight, 0, 0, OrWidth, OrHeight, UnitPixel, 0, 0, 0
    Else
        GdipDrawImageRectI Graphics, Img, 0, 0, Width, Height
    End If
    GdipDeleteGraphics Graphics
End Sub

' Replaces the old bitmap of the hDC, Returns the bitmap and Deletes the hDC
Private Sub GetBitmap(hdc As Long, hBitmap As Long)
    hBitmap = SelectObject(hdc, hBitmap)
    DeleteDC hdc
End Sub

' Creates a Picture Object from a handle to a bitmap
Public Function CreatePicture(hBitmap As Long) As IPicture
    Dim IID_IDispatch As GUID
    Dim pic           As PICTDESC
    Dim IPic          As IPicture
    
    ' Fill in OLE IDispatch Interface ID
    IID_IDispatch.Data1 = &H20400
    IID_IDispatch.Data4(0) = &HC0
    IID_IDispatch.Data4(7) = &H46
        
    ' Fill Pic with necessary parts
    pic.Size = Len(pic)        ' Length of structure
    pic.Type = PICTYPE_BITMAP  ' Type of Picture (bitmap)
    pic.hBmp = hBitmap         ' Handle to bitmap

    ' Create the picture
    OleCreatePictureIndirect pic, IID_IDispatch, True, IPic
    Set CreatePicture = IPic
End Function

' Returns a resized version of the picture
Public Function ResizeGDIPlus(Handle As Long, _
                              PicType As PictureTypeConstants, _
                              ByVal Width As Long, ByVal Height As Long, _
                              Optional BackColor As Long = vbWhite, _
                              Optional RetainRatio As Boolean = False) As IPicture
    Dim Img       As Long
    Dim hdc       As Long
    Dim hBitmap   As Long
    Dim WmfHeader As wmfPlaceableFileHeader
   
     ' Initialise GDI+
    Token = InitGDIPlus
    
    If Token <> 0 Then
    
    ' Determine pictyre type
    Select Case PicType
    Case vbPicTypeBitmap
         GdipCreateBitmapFromHBITMAP Handle, ByVal 0&, Img
    Case vbPicTypeMetafile
         FillInWmfHeader WmfHeader, Width, Height
         GdipCreateMetafileFromWmf Handle, False, WmfHeader, Img
    Case vbPicTypeEMetafile
         GdipCreateMetafileFromEmf Handle, False, Img
    Case vbPicTypeIcon
         ' Does not return a valid Image object
         GdipCreateBitmapFromHICON Handle, Img
    End Select
    
    ' Continue with resizing only if we have a valid image object
    If Img Then
        InitDC hdc, hBitmap, BackColor, Width, Height
        gdipResize Img, hdc, Width, Height, RetainRatio
        GdipDisposeImage Img
        GetBitmap hdc, hBitmap
        Set ResizeGDIPlus = CreatePicture(hBitmap)
    End If
    ' Free GDI+
    FreeGDIPlus Token
    Else
        MsgBox " Error " + Str(Token) + " : " + GdiErrorString(Token), vbCritical
    End If
End Function

' Fills in the wmfPlacable header
Private Sub FillInWmfHeader(WmfHeader As wmfPlaceableFileHeader, Width As Long, Height As Long)
    WmfHeader.BoundingBox.Right = Width
    WmfHeader.BoundingBox.Bottom = Height
    WmfHeader.Inch = 1440
    WmfHeader.Key = GDIP_WMF_PLACEABLEKEY
End Sub

Public Function GdiErrorString(ByVal lError As Long) As String
  Dim s As String
  
  Select Case lError
    Case 1: s = "Generic Error"
    Case 2: s = "Invalid Parameter"
    Case 3: s = "Out Of Memory"
    Case 4: s = "Object Busy"
    Case 5: s = "Insufficient Buffer"
    Case 6: s = "Not Implemented"
    Case 7: s = "Win32 Error"
    Case 8: s = "Wrong State"
    Case 9: s = "Aborted"
    Case 10: s = "File Not Found"
    Case 11: s = "Value Overflow"
    Case 12: s = "Access Denied"
    Case 13: s = "Unknown Image Format"
    Case 14: s = "FontFamily Not Found"
    Case 15: s = "FontStyle Not Found"
    Case 16: s = "Not TrueType Font"
    Case 17: s = "Unsupported Gdiplus Version"
    Case 18: s = "Gdiplus Not Initialized"
    Case 19: s = "Property Not Found"
    Case 20: s = "Property Not Supported"
    Case Else: s = "Unknown GDI+ Error"
  End Select
  
  GdiErrorString = s
End Function
