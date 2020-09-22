Attribute VB_Name = "SaveBitmapAs"
' //////////////////////////////////
' // Bitmap to PNG, GIF, TIF, JPG //
' //////////////////////////////////
'
' I found this on MS newsgroups awhile back,
' it uses GDI+.
'
' I modified it slightly so it auto selects,
' the file type depending on the file extension
' you use (.PNG, .GIF, .TIF .JPG, .BMP).
'
'hth,
'Edgemeal
'
Option Explicit

Private GdipToken       As Long
Private GdipInitialized As Boolean

Public Const GdiPlusVersion     As Long = 1
Private Const CP_ACP            As Long = 0

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type ImageCodecInfo
   ClassID As GUID
   FormatID As GUID
   CodecName As Long
   DllName As Long
   FormatDescription As Long
   FilenameExtension As Long
   MimeType As Long
   Flags As Long
   Version As Long
   SigCount As Long
   SigSize As Long
   SigPattern As Long
   SigMask As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Enum SmoothingMode
    SmoothingModeInvalid = -1&
    SmoothingModeDefault = 0&
    SmoothingModeHighSpeed = 1&
    SmoothingModeHighQuality = 2&
    SmoothingModeNone = 3&
    SmoothingModeAntiAlias8x4 = 4&
    SmoothingModeAntiAlias = SmoothingModeAntiAlias8x4
    'SmoothingModeAntiAlias8x8
End Enum

' ----==== GDIPlus Enums ====----
Private Enum GpStatus 'GDI+ Status
    OK = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codepage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal Filename As Long, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hbm As Long, ByVal hpal As Long, nBitmap As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (ByRef numEncoders As Long, ByRef Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, ByRef Encoders As Any) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal SmoothingMode As SmoothingMode) As GpStatus
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef Graphics As Long) As Long
Private Declare Function GdipSaveGraphics Lib "gdiplus" (ByVal Graphics As Long, STATE As Long) As GpStatus

Public Function SavePictureFromHDC(ByVal Obj As Object, ByVal sFileName As String) As Boolean
    Dim lBitmap As Long
    Dim PicEncoder As GUID
    Dim sID As String
    Dim Graphics As Long
    Dim hBitmap As Long
    Dim STATE As Long
    ' Use file name extention to determine,
    ' what format we want to save the file in.
    
    Select Case LCase$(Right$(sFileName, 4))
        Case ".png"
            sID = "image/png"
        Case ".gif"
            sID = "image/gif"
        Case ".jpg", ".jpeg"
            sID = "image/jpeg"
        Case ".tif", ".tiff"
            sID = "image/tiff"
        Case ".bmp"
            sID = "image/bmp"
        Case ".emf"
           sID = "image/x-emf"
        Case ".wmf"
           sID = "image/x-wmf"
        Case Else
            Exit Function
    End Select
    
    Call GdipCreateFromHDC(Obj.hDC, Graphics)
    Call GdipSetSmoothingMode(Graphics, SmoothingModeAntiAlias)
    Call GdipSaveGraphics(Graphics, STATE)
        hBitmap = Obj.Picture

    If GdipCreateBitmapFromHBITMAP(hBitmap, 0&, lBitmap) = 0 Then
    
        If GetEncoderClsid(sID, PicEncoder) = True Then
            SavePictureFromHDC = (GdipSaveImageToFile(lBitmap, StrPtr(sFileName), PicEncoder, ByVal 0) = 0)
        End If
        GdipDisposeImage lBitmap
    End If
End Function

Private Function GetEncoderClsid(strMimeType As String, ClassID As GUID) As Boolean
    
    Dim num As Long
    Dim Size As Long
    Dim imgCodecInfo() As ImageCodecInfo
    Dim lval As Long
    Dim buffer() As Byte

    GdipGetImageEncodersSize num, Size
    If Size Then
        ReDim imgCodecInfo(num) As ImageCodecInfo
        ReDim buffer(Size) As Byte

        GdipGetImageEncoders num, Size, buffer(0)
        CopyMemory imgCodecInfo(0), buffer(0), (Len(imgCodecInfo(0)) * num)

        For lval = 0 To num - 1
            'image/bmp,image/jpeg,image/gif,image/tiff,image/png
            If StrComp(GetStrFromPtrW(imgCodecInfo(lval).MimeType), strMimeType, vbTextCompare) = 0 Then
                ClassID = imgCodecInfo(lval).ClassID
                GetEncoderClsid = True
                Exit For
            End If
        Next
        Erase imgCodecInfo
        Erase buffer
    End If
    
End Function

Private Function GetStrFromPtrW(lpszW As Long) As String
    
    Dim sRV As String

    sRV = String$(lstrlenW(ByVal lpszW) * 2, vbNullChar)
    WideCharToMultiByte CP_ACP, 0, ByVal lpszW, -1, ByVal sRV, Len(sRV), 0, 0
    GetStrFromPtrW = Left$(sRV, lstrlenW(StrPtr(sRV)))
    
End Function

Public Function StartUpGDIPlus(ByVal GdipVersion As Long) As Boolean
    
    Dim GdipStartupInput As GDIPlusStartupInput
    
    GdipStartupInput.GdiPlusVersion = GdipVersion
    GdipInitialized = (GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0&) = 0)
    StartUpGDIPlus = GdipInitialized
End Function

Public Sub ShutdownGDIPlus()
    
    If GdipInitialized Then
        GdiplusShutdown GdipToken
    End If
    
End Sub


