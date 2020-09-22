Attribute VB_Name = "mFilterG"
'+--------------------------------------------------------+
'| Name            : mFilterG - Graphic Filters           |
'| Author          : Manuel Augusto Nogueira dos Santos   |
'|                  & Diomidis Kiriakopoulos + 17 filtes  |
'| Dates           : 23/03/2001 - 20/03/2009              |
'| Description     : Apply effects to images              |
'+--------------------------------------------------------+
'| FilterG(ByVal Filtro As iFilterG                       |
'|               > one of iFilterG Enum                   |
'|         ByVal Pic As Long,                             |
'|               > PictureBox.Image                       |
'|         ByVal Factor As Long,                          |
'|               > depends upon Filtro (see below)        |
'|         ByRef pProgress As Long)                       |
'|               > % progress done                        |
'+--------------------------------------------------------+
'| Factor                                                 |
'|  iSHARPEN    : 0..N for Sharpen + to Sharpen -         |
'|  iNEGATIVE   : no effect                               |
'|  iBLUR       : no effect                               |
'|  iDIFFUSE    : diffuse radius, 6 normal / 12 diffuse + |
'|  iSMOOTH     : no effect                               |
'|  iEDGE       : 1..N for EdgeEnhance + to EdgeEnhance - |
'|  iCONTOUR    : RGB BackColor                           |
'|  iEMBOSS     : RGB BackColor                           |
'|  iEMBOSSMORE : RGB BackColor                           |
'|  iENGRAVE    : RGB BackColor                           |
'|  iENGRAVEMORE: RGB BackColor                           |
'|  iGREYSCALE  : no effect                               |
'|  iRELIEF     : no effect                               |
'|  iBRIGHTNESS : >0 to increase, <0 to decrease          |
'|  iPIXELIZE   : size of each pixel                      |
'|  iSWAPBANK   : 1..5 RGB to (BRG,GBR,RBG,BGR,GRB)       |
'|  iCONTRAST   : >0 to increase, <0 to decrease          |
'|  iCOLDEPTH1  : RGB color to set black below            |
'|  iCOLDEPTH2  : no effect                               |
'|  iCOLDEPTH3  : no effect                               |
'|  iCOLDEPTH4  : 1..n Palette colors weight              |
'|  iCOLDEPTH5  : 1..n Palette colors weight              |
'|  iCOLDEPTH6  : 1..n Palette colors weight              |
'|  iAQUA       : no effect                               |
'|  iDILATE     : no effect                               |
'|  iERODE      : no effect                               |
'|  iCONNECTION : no effect                               |
'|  iSTRETCH    : no effect                               |
'|  iADDNOISE   : noise intensity                         |
'|  iSATURATION : >0 to increase, <0 to decrease          |

'|  iNEON       : no effect                               |
'|  iGAMMA      : 1-100                                   |
'|  iGrid3d     : 1..n                                    |
'|  iMirrorRL   : no effect                               |
'|  iMirrorLR   : no effect                               |
'|  iMirrorDT   : no effect                               |
'|  iMirrorTD   : no effect                               |
'|  iArt        : no effect                               |
'|  iStranges   : 1..n                                    |
'|  iFog        : 1..n                                    |
'|  iSnow       : 4-64                                    |
'|  iWave       : 0 - 16                                  |
'|  iCrease     : >=64- 65536<  default:512               |
'|  iSepia      : no effect                               |
'|  iRects:     : no effect                               |
'|  iComic      : no effect                               |
'|  iIce:       : no effect                               |
'+--------------------------------------------------------+
Option Explicit

'---------------Public var
Public WorkFilterG As Boolean

Public Enum iFilterG
    iSharpen = 1
    iNegative = 2
    iBlur = 3
    iDiffuse = 4
    iSmooth = 5
    iEDGE = 6
    iContour = 7
    iEmboss = 8
    iEmbossMore = 9
    iEngrave = 10
    iEngraveMore = 11
    iGreyScale = 12
    iRelief = 13
    iBRIGHTNESS = 14
    iPixelize = 15
    iSwapBank = 16
    iContrast = 17
    iColDepth1 = 18
    iColDepth2 = 19
    iColDepth3 = 20
    iColDepth4 = 21
    iColDepth5 = 22
    iColDepth6 = 23
    iAqua = 24
    iDilate = 25
    iErode = 26
    iConnection = 27
    iStretch = 28
    iAddNoise = 29
    iSaturation = 30
    iGamma = 31
    iNeon = 32
    iGrid3d = 33
    iMirrorRL = 34
    iMirrorLR = 35
    iMirrorDT = 36
    iMirrorTD = 37
    iArt = 38
    iStranges = 39
    iFog = 40
    iSnow = 41
    iWave = 42
    iCrease = 43
    iSepia = 44
    iRects = 45 'dif
    iComic = 46 'color
    iIce = 47  'color
End Enum

'--------------------Private var
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0&
Private Const LR_LOADFROMFILE = &H10
Private Const IMAGE_BITMAP = 0&

Private iDATA() As Byte           'holds bitmap data
Private bDATA() As Byte           'holds bitmap backup
Private PicInfo As BITMAP         'bitmap info structure
Private DIBInfo As BITMAPINFO     'Device Ind. Bitmap info structure
Private mProgress As Long         '% filter progress
Private Speed(0 To 765) As Long   'Speed up values
Public mCancel As Boolean        'Cancel Progress

Public Function FilterG(ByVal Filtro As iFilterG, ByVal pic As Long, ByVal Factor As Long, ByRef pProgress As Long) As Boolean
  Dim hdcNew As Long
  Dim oldhand As Long
  Dim ret As Long
  Dim BytesPerScanLine As Long
  Dim PadBytesPerScanLine As Long
    
  If WorkFilterG = True Then Exit Function
  WorkFilterG = True
  mCancel = False
  
  On Error GoTo FilterError:
  'get data buffer
  Call GetObjectAPI(pic, Len(PicInfo), PicInfo)
  hdcNew = CreateCompatibleDC(0&)
  oldhand = SelectObject(hdcNew, pic)
  With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = PicInfo.bmWidth
    .biHeight = -PicInfo.bmHeight     'bottom up scan line is now inverted
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    PadBytesPerScanLine = BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With
  'redimension  (BGR+pad,x,y)
  ReDim iDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
  ReDim bDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
  'get bytes
  ret = GetDIBits(hdcNew, pic, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
  ret = GetDIBits(hdcNew, pic, 0, PicInfo.bmHeight, bDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
  'do it
  Select Case Filtro
    Case iSharpen:          Call Sharpen(pProgress, Factor)
    Case iNegative:         Call NegativeImage(pProgress)
    Case iBlur:             Call Blurs(pProgress)
    Case iDiffuse:          Call Diffuse(pProgress, Factor)
    Case iSmooth:           Call Smooth(pProgress)
    Case iEDGE:             Call EdgeEnhance(pProgress, Factor)
    Case iContour:          Call Contour(pProgress, Factor)
    Case iEmboss:           Call Emboss(pProgress, Factor)
    Case iEmbossMore:       Call EmbossMore(pProgress, Factor)
    Case iEngrave:          Call Engrave(pProgress, Factor)
    Case iEngraveMore:      Call EngraveMore(pProgress, Factor)
    Case iGreyScale:        Call GreyScale(pProgress)
    Case iRelief:           Call Relief(pProgress)
    Case iBRIGHTNESS:       Call Brightness(pProgress, Factor)
    Case iPixelize:         Call Pixelize(pProgress, Factor)
    Case iSwapBank:         Call SwapBank(pProgress, Factor)
    Case iContrast:         Call Contrast(pProgress, Factor)
    Case iColDepth1:        Call NearestColorBW(pProgress, Factor)
    Case iColDepth2:        Call EnhancedDiffusionBW(pProgress)
    Case iColDepth3:        Call OrderedDitherBW(pProgress)
    Case iColDepth4:        Call FloydSteinbergBW(pProgress, Factor)
    Case iColDepth5:        Call BurkeBW(pProgress, Factor)
    Case iColDepth6:        Call StuckiBW(pProgress, Factor)
    Case iAqua:             Call Aqua(pProgress)
    Case iDilate:           Call Dilate(pProgress)
    Case iErode:            Call Erode(pProgress)
    Case iConnection:       Call ConnectedContour(pProgress)
    Case iStretch:          Call StretchHistogram(pProgress)
    Case iAddNoise:         Call AddNoise(pProgress, Factor)
    Case iSaturation:       Call Saturation(pProgress, Factor)
    Case iGamma:            Call GammaCorrection(pProgress, Factor)
    Case iNeon:             Call Neon(pProgress)
    Case iGrid3d:           Call Grid3d(pProgress, Factor)
    Case iMirrorRL:         Call MirrorRightLeft(pProgress)
    Case iMirrorLR:         Call MirrorLeftRight(pProgress)
    Case iMirrorDT:         Call MirrorDownTop(pProgress)
    Case iMirrorTD:         Call MirrorTopDown(pProgress)
    Case iArt:              Call Art(pProgress)
    Case iStranges:         Call Stranges(pProgress, Factor)
    Case iFog:              Call Fog(pProgress, Factor)
    Case iSnow:             Call Snow(pProgress, Factor)
    Case iWave:             Call Wave(pProgress, Factor) '0-16
    Case iCrease:           Call Wave(pProgress, Factor) '64-65536 >512
    Case iSepia:            Call Sepia(pProgress)
    Case iRects:            Call Rects(pProgress)
    Case iComic:            Call Comic(pProgress)
    Case iIce:              Call Ice(pProgress)
  End Select
  
  'copy bytes to device
  ret = SetDIBits(hdcNew, pic, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
  SelectObject hdcNew, oldhand
  DeleteDC hdcNew
  ReDim iDATA(1 To 4, 1 To 2, 1 To 2) As Byte
  ReDim bDATA(1 To 4, 1 To 2, 1 To 2) As Byte
  WorkFilterG = False
  FilterG = mCancel
  Exit Function
FilterError:
  MsgBox "Filter Error"
  WorkFilterG = False
End Function

'-------------------------------------------AUXILIARY
Private Sub GetRGB(ByVal Col As Long, ByRef r As Long, ByRef g As Long, ByRef b As Long)
  r = Col Mod 256
  g = ((Col And &HFF00&) \ 256&) Mod 256&
  b = (Col And &HFF0000) \ 65536
End Sub

'-------------------------------------------FILTERS

Private Sub NegativeImage(ByRef pProgress As Long)
  Dim X As Long, Y As Long
  
  mProgress = 0
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      iDATA(1, X, Y) = 255 - iDATA(1, X, Y)
      iDATA(2, X, Y) = 255 - iDATA(2, X, Y)
      iDATA(3, X, Y) = 255 - iDATA(3, X, Y)
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Sharpen(ByRef pProgress As Long, ByVal Factor As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim mf As Long, dF As Long
  On Error Resume Next
  mProgress = 0
  mf = 24 + Factor
  dF = 8 + Factor
  For Y = 2 To PicInfo.bmHeight - 1
    For X = 2 To PicInfo.bmWidth - 1
      b = CLng(iDATA(1, X, Y - 1)) + CLng(iDATA(1, X - 1, Y)) + CLng(iDATA(1, X + 1, Y)) + CLng(iDATA(1, X, Y + 1)) + _
          CLng(iDATA(1, X + 1, Y + 1)) + CLng(iDATA(1, X - 1, Y + 1)) + CLng(iDATA(1, X + 1, Y - 1)) + CLng(iDATA(1, X - 1, Y - 1))
      b = (mf * CLng(iDATA(1, X, Y)) - 2 * b) \ dF
      
      g = CLng(iDATA(2, X, Y - 1)) + CLng(iDATA(2, X - 1, Y)) + CLng(iDATA(2, X + 1, Y)) + CLng(iDATA(2, X, Y + 1)) + _
          CLng(iDATA(2, X + 1, Y + 1)) + CLng(iDATA(2, X - 1, Y + 1)) + CLng(iDATA(2, X + 1, Y - 1)) + CLng(iDATA(2, X - 1, Y - 1))
      
      g = (mf * CLng(iDATA(2, X, Y)) - 2 * g) \ dF
      
      r = CLng(iDATA(3, X, Y - 1)) + CLng(iDATA(3, X - 1, Y)) + CLng(iDATA(3, X + 1, Y)) + CLng(iDATA(3, X, Y + 1)) + _
          CLng(iDATA(3, X + 1, Y + 1)) + CLng(iDATA(3, X - 1, Y + 1)) + CLng(iDATA(3, X + 1, Y - 1)) + CLng(iDATA(3, X - 1, Y - 1))
      r = (mf * CLng(iDATA(3, X, Y)) - 2 * r) \ dF
      
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If g > 255 Then g = 255
      If g < 0 Then g = 0
      If b > 255 Then b = 255
      If b < 0 Then b = 0
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Engrave(ByRef pProgress As Long, ByVal BackCol As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim cB As Long, cG As Long, cR As Long
   On Error Resume Next
  mProgress = 0
  Call GetRGB(BackCol, cR, cG, cB)
  For Y = 1 To PicInfo.bmHeight '- 1
    For X = 1 To PicInfo.bmWidth '- 1
      b = Abs(CLng(iDATA(1, X + 1, Y + 1)) - CLng(iDATA(1, X, Y)) + cB)
      g = Abs(CLng(iDATA(2, X + 1, Y + 1)) - CLng(iDATA(2, X, Y)) + cG)
      r = Abs(CLng(iDATA(3, X + 1, Y + 1)) - CLng(iDATA(3, X, Y)) + cR)
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If g > 255 Then g = 255
      If g < 0 Then g = 0
      If b > 255 Then b = 255
      If b < 0 Then b = 0
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub EngraveMore(ByRef pProgress As Long, ByVal BackCol As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim cB As Long, cG As Long, cR As Long
  On Error Resume Next
  mProgress = 0
  Call GetRGB(BackCol, cR, cG, cB)
  For Y = 1 To PicInfo.bmHeight '- 1
    For X = 1 To PicInfo.bmWidth '- 1
      b = CLng(bDATA(1, X + 1, Y - 1)) - CLng(bDATA(1, X - 1, Y - 1)) + _
          CLng(bDATA(1, X + 1, Y)) - CLng(bDATA(1, X - 1, Y)) + _
          CLng(bDATA(1, X + 1, Y + 1)) - CLng(bDATA(1, X - 1, Y + 1)) + cB
      g = CLng(bDATA(2, X + 1, Y - 1)) - CLng(bDATA(2, X - 1, Y - 1)) + _
          CLng(bDATA(2, X + 1, Y)) - CLng(bDATA(2, X - 1, Y)) + _
          CLng(bDATA(2, X + 1, Y + 1)) - CLng(bDATA(2, X - 1, Y + 1)) + cG
      r = CLng(bDATA(3, X + 1, Y - 1)) - CLng(bDATA(3, X - 1, Y - 1)) + _
          CLng(bDATA(3, X + 1, Y)) - CLng(bDATA(3, X - 1, Y)) + _
          CLng(bDATA(3, X + 1, Y + 1)) - CLng(bDATA(3, X - 1, Y + 1)) + cR
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If g > 255 Then g = 255
      If g < 0 Then g = 0
      If b > 255 Then b = 255
      If b < 0 Then b = 0
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Emboss(ByRef pProgress As Long, ByVal BackCol As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim cB As Long, cG As Long, cR As Long
  On Error Resume Next
  mProgress = 0
  Call GetRGB(BackCol, cR, cG, cB)
  For Y = 1 To PicInfo.bmHeight '- 1
    For X = 1 To PicInfo.bmWidth '- 1
      b = Abs(CLng(iDATA(1, X, Y)) - CLng(iDATA(1, X + 1, Y + 1)) + cB)
      g = Abs(CLng(iDATA(2, X, Y)) - CLng(iDATA(2, X + 1, Y + 1)) + cG)
      r = Abs(CLng(iDATA(3, X, Y)) - CLng(iDATA(3, X + 1, Y + 1)) + cR)
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If g > 255 Then g = 255
      If g < 0 Then g = 0
      If b > 255 Then b = 255
      If b < 0 Then b = 0
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub EmbossMore(ByRef pProgress As Long, ByVal BackCol As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim cB As Long, cG As Long, cR As Long
   On Error Resume Next
  mProgress = 0
  Call GetRGB(BackCol, cR, cG, cB)
  For Y = 2 To PicInfo.bmHeight '- 1
    For X = 2 To PicInfo.bmWidth '- 1
      b = CLng(bDATA(1, X - 1, Y - 1)) - CLng(bDATA(1, X + 1, Y - 1)) + _
          CLng(bDATA(1, X - 1, Y)) - CLng(bDATA(1, X + 1, Y)) + _
          CLng(bDATA(1, X - 1, Y + 1)) - CLng(bDATA(1, X + 1, Y + 1)) + cB
      g = CLng(bDATA(2, X - 1, Y - 1)) - CLng(bDATA(2, X + 1, Y - 1)) + _
          CLng(bDATA(2, X - 1, Y)) - CLng(bDATA(2, X + 1, Y)) + _
          CLng(bDATA(2, X - 1, Y + 1)) - CLng(bDATA(2, X + 1, Y + 1)) + cG
      r = CLng(bDATA(3, X - 1, Y - 1)) - CLng(bDATA(3, X + 1, Y - 1)) + _
          CLng(bDATA(3, X - 1, Y)) - CLng(bDATA(3, X + 1, Y)) + _
          CLng(bDATA(3, X - 1, Y + 1)) - CLng(bDATA(3, X + 1, Y + 1)) + cR
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If g > 255 Then g = 255
      If g < 0 Then g = 0
      If b > 255 Then b = 255
      If b < 0 Then b = 0
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Diffuse(ByRef pProgress As Long, ByVal Factor As Long)
  Dim X As Long, Y As Long
  Dim aX As Long, aY As Long
  Dim r As Long, g As Long, b As Long
  Dim hF As Long

  mProgress = 0
  hF = Factor / 2
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      aX = Rnd * Factor - hF
      aY = Rnd * Factor - hF
      If X + aX < 1 Then aX = 0
      If X + aX > PicInfo.bmWidth Then aX = 0
      If Y + aY < 1 Then aY = 0
      If Y + aY > PicInfo.bmHeight Then aY = 0
      iDATA(1, X, Y) = iDATA(1, X + aX, Y + aY)
      iDATA(2, X, Y) = iDATA(2, X + aX, Y + aY)
      iDATA(3, X, Y) = iDATA(3, X + aX, Y + aY)
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Smooth(ByRef pProgress As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  On Error Resume Next
  mProgress = 0
  For Y = 2 To PicInfo.bmHeight - 1
    For X = 2 To PicInfo.bmWidth - 1
      b = CLng(iDATA(1, X, Y)) + _
        CLng(iDATA(1, X - 1, Y)) + CLng(iDATA(1, X, Y - 1)) + _
        CLng(iDATA(1, X, Y + 1)) + CLng(iDATA(1, X + 1, Y))
      b = b \ 5
      g = CLng(iDATA(2, X, Y)) + _
        CLng(iDATA(2, X - 1, Y)) + CLng(iDATA(2, X, Y - 1)) + _
        CLng(iDATA(2, X, Y + 1)) + CLng(iDATA(2, X + 1, Y))
      g = g \ 5
      r = CLng(iDATA(3, X, Y)) + _
        CLng(iDATA(3, X - 1, Y)) + CLng(iDATA(3, X, Y - 1)) + _
        CLng(iDATA(3, X, Y + 1)) + CLng(iDATA(3, X + 1, Y))
      r = r \ 5
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Blurs(ByRef pProgress As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  On Error Resume Next
  mProgress = 0
  For Y = 2 To PicInfo.bmHeight - 1
    For X = 2 To PicInfo.bmWidth - 1
      b = CLng(iDATA(1, X - 1, Y - 1)) + CLng(iDATA(1, X - 1, Y)) + _
          CLng(iDATA(1, X - 1, Y + 1)) + CLng(iDATA(1, X, Y - 1)) + _
          CLng(iDATA(1, X, Y + 1)) + CLng(iDATA(1, X + 1, Y - 1)) + _
          CLng(iDATA(1, X + 1, Y)) + CLng(iDATA(1, X + 1, Y + 1))
      b = b \ 8
      g = CLng(iDATA(2, X - 1, Y - 1)) + CLng(iDATA(2, X - 1, Y)) + _
          CLng(iDATA(2, X - 1, Y + 1)) + CLng(iDATA(2, X, Y - 1)) + _
          CLng(iDATA(2, X, Y + 1)) + CLng(iDATA(2, X + 1, Y - 1)) + _
          CLng(iDATA(2, X + 1, Y)) + CLng(iDATA(2, X + 1, Y + 1))
      g = g \ 8
      r = CLng(iDATA(3, X - 1, Y - 1)) + CLng(iDATA(3, X - 1, Y)) + _
          CLng(iDATA(3, X - 1, Y + 1)) + CLng(iDATA(3, X, Y - 1)) + _
          CLng(iDATA(3, X, Y + 1)) + CLng(iDATA(3, X + 1, Y - 1)) + _
          CLng(iDATA(3, X + 1, Y)) + CLng(iDATA(3, X + 1, Y + 1))
      r = r \ 8
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub EdgeEnhance(ByRef pProgress As Long, ByVal Factor As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim mf As Long, dF As Long
  On Error Resume Next
  mProgress = 0
  mf = 9 + Factor
  dF = 1 + Factor
  For Y = 1 To PicInfo.bmHeight '- 1
    For X = 1 To PicInfo.bmWidth '- 1
      b = CLng(iDATA(1, X - 1, Y - 1)) + CLng(iDATA(1, X - 1, Y)) + _
        CLng(iDATA(1, X - 1, Y + 1)) + CLng(iDATA(1, X, Y - 1)) + _
        CLng(iDATA(1, X, Y + 1)) + CLng(iDATA(1, X + 1, Y - 1)) + _
        CLng(iDATA(1, X + 1, Y)) + CLng(iDATA(1, X + 1, Y + 1))
      b = (mf * CLng(iDATA(1, X, Y)) - b) \ dF
      g = CLng(iDATA(2, X - 1, Y - 1)) + CLng(iDATA(2, X - 1, Y)) + _
        CLng(iDATA(2, X - 1, Y + 1)) + CLng(iDATA(2, X, Y - 1)) + _
        CLng(iDATA(2, X, Y + 1)) + CLng(iDATA(2, X + 1, Y - 1)) + _
        CLng(iDATA(2, X + 1, Y)) + CLng(iDATA(2, X + 1, Y + 1))
      g = (mf * CLng(iDATA(2, X, Y)) - g) \ dF
      r = CLng(iDATA(3, X - 1, Y - 1)) + CLng(iDATA(3, X - 1, Y)) + _
        CLng(iDATA(3, X - 1, Y + 1)) + CLng(iDATA(3, X, Y - 1)) + _
        CLng(iDATA(3, X, Y + 1)) + CLng(iDATA(3, X + 1, Y - 1)) + _
        CLng(iDATA(3, X + 1, Y)) + CLng(iDATA(3, X + 1, Y + 1))
      r = (mf * CLng(iDATA(3, X, Y)) - r) \ dF
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If g > 255 Then g = 255
      If g < 0 Then g = 0
      If b > 255 Then b = 255
      If b < 0 Then b = 0
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Contour(ByRef pProgress As Long, ByVal BackCol As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim cB As Long, cG As Long, cR As Long
  On Error Resume Next
  mProgress = 0
  Call GetRGB(BackCol, cR, cG, cB)
  For Y = 1 To PicInfo.bmHeight '- 1
    For X = 1 To PicInfo.bmWidth '- 1
      b = CLng(bDATA(1, X - 1, Y - 1)) + CLng(bDATA(1, X - 1, Y)) + _
          CLng(bDATA(1, X - 1, Y + 1)) + CLng(bDATA(1, X, Y - 1)) + _
          CLng(bDATA(1, X, Y + 1)) + CLng(bDATA(1, X + 1, Y - 1)) + _
          CLng(bDATA(1, X + 1, Y)) + CLng(bDATA(1, X + 1, Y + 1))
      g = CLng(bDATA(2, X - 1, Y - 1)) + CLng(bDATA(2, X - 1, Y)) + _
          CLng(bDATA(2, X - 1, Y + 1)) + CLng(bDATA(2, X, Y - 1)) + _
          CLng(bDATA(2, X, Y + 1)) + CLng(bDATA(2, X + 1, Y - 1)) + _
          CLng(bDATA(2, X + 1, Y)) + CLng(bDATA(2, X + 1, Y + 1))
      r = CLng(bDATA(3, X - 1, Y - 1)) + CLng(bDATA(3, X - 1, Y)) + _
          CLng(bDATA(3, X - 1, Y + 1)) + CLng(bDATA(3, X, Y - 1)) + _
          CLng(bDATA(3, X, Y + 1)) + CLng(bDATA(3, X + 1, Y - 1)) + _
          CLng(bDATA(3, X + 1, Y)) + CLng(bDATA(3, X + 1, Y + 1))
      b = 8 * CLng(bDATA(1, X, Y)) - b + cB
      g = 8 * CLng(bDATA(2, X, Y)) - g + cG
      r = 8 * CLng(bDATA(3, X, Y)) - r + cR
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If g > 255 Then g = 255
      If g < 0 Then g = 0
      If b > 255 Then b = 255
      If b < 0 Then b = 0
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub GreyScale(ByRef pProgress As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  
  mProgress = 0
  For X = 0 To 765
    Speed(X) = X \ 3
  Next X
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = iDATA(1, X, Y)
      g = iDATA(2, X, Y)
      r = iDATA(3, X, Y)
      b = Speed(r + g + b)
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = b
      iDATA(3, X, Y) = b
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Relief(ByRef pProgress As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  On Error Resume Next
  mProgress = 0
  For Y = 1 To PicInfo.bmHeight '- 1
    For X = 1 To PicInfo.bmWidth '- 1
      b = 2 * CLng(bDATA(1, X - 1, Y - 1)) + CLng(bDATA(1, X - 1, Y)) + _
          CLng(bDATA(1, X, Y - 1)) - CLng(bDATA(1, X, Y + 1)) - _
          CLng(bDATA(1, X + 1, Y)) - 2 * CLng(bDATA(1, X + 1, Y + 1))
      g = 2 * CLng(bDATA(2, X - 1, Y - 1)) + CLng(bDATA(2, X - 1, Y)) + _
          CLng(bDATA(2, X, Y - 1)) - CLng(bDATA(2, X, Y + 1)) - _
          CLng(bDATA(2, X + 1, Y)) - 2 * CLng(bDATA(2, X + 1, Y + 1))
      r = 2 * CLng(bDATA(3, X - 1, Y - 1)) + CLng(bDATA(3, X - 1, Y)) + _
          CLng(bDATA(3, X, Y - 1)) - CLng(bDATA(3, X, Y + 1)) - _
          CLng(bDATA(3, X + 1, Y)) - 2 * CLng(bDATA(3, X + 1, Y + 1))
      b = (CLng(bDATA(1, X, Y)) + b) \ 2 + 50
      g = (CLng(bDATA(2, X, Y)) + g) \ 2 + 50
      r = (CLng(bDATA(3, X, Y)) + r) \ 2 + 50
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If g > 255 Then g = 255
      If g < 0 Then g = 0
      If b > 255 Then b = 255
      If b < 0 Then b = 0
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Brightness(ByRef pProgress As Long, ByVal Factor As Long)
  Dim X As Long, Y As Long
  Dim sF As Single
  
  mProgress = 0
  sF = (Factor + 100) / 100
  For X = 0 To 255
    Speed(X) = X * sF
    If Speed(X) > 255 Then Speed(X) = 255
    If Speed(X) < 0 Then Speed(X) = 0
  Next X
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      iDATA(1, X, Y) = Speed(bDATA(1, X, Y))
      iDATA(2, X, Y) = Speed(bDATA(2, X, Y))
      iDATA(3, X, Y) = Speed(bDATA(3, X, Y))
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Contrast(ByRef pProgress As Long, ByVal Factor As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim sF As Single
  Dim mCol As Long, nCol As Long

  mProgress = 0
  For X = 0 To 765
    Speed(X) = X \ 3
  Next X
  mCol = 0
  nCol = 0
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      mCol = mCol + Speed(r + g + b)
      nCol = nCol + 1
    Next X
  Next Y
  mCol = mCol \ nCol
  sF = (Factor + 100) / 100
  For X = 0 To 255
    Speed(X) = (X - mCol) * sF + mCol
  Next X
  pProgress = 5
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = Speed(CLng(bDATA(1, X, Y)))
      g = Speed(CLng(bDATA(2, X, Y)))
      r = Speed(CLng(bDATA(3, X, Y)))
      Do While (b < 0) Or (b > 255) Or (g < 0) Or (g > 255) Or (r < 0) Or (r > 255)
        If (b <= 0) And (g <= 0) And (r <= 0) Then
          b = 0
          g = 0
          r = 0
        End If
        If (b >= 255) And (g >= 255) And (r >= 255) Then
          b = 255
          g = 255
          r = 255
        End If
        If b < 0 Then
          g = g + b \ 2
          r = r + b \ 2
          b = 0
        End If
        If b > 255 Then
          g = g + (b - 255) \ 2
          r = r + (b - 255) \ 2
          b = 255
        End If
        If g < 0 Then
          b = b + g \ 2
          r = r + g \ 2
          g = 0
        End If
        If g > 255 Then
          b = b + (g - 255) \ 2
          r = r + (g - 255) \ 2
          g = 255
        End If
        If r < 0 Then
          g = g + r \ 2
          b = b + r \ 2
          r = 0
        End If
        If r > 255 Then
          g = g + (r - 255) \ 2
          b = b + (r - 255) \ 2
          r = 255
        End If
      Loop
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = 5 + (Y * 95) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Pixelize(ByRef pProgress As Long, ByVal PixSize As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim pX As Long, pY As Long
  Dim sx As Long, sy As Long
  Dim mC As Long
  
  mProgress = 0
  b = 0: g = 0: r = 0
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      If ((X - 1) Mod PixSize) = 0 Then
        sx = ((X - 1) \ PixSize) * PixSize + 1
        sy = ((Y - 1) \ PixSize) * PixSize + 1
        b = 0: g = 0: r = 0: mC = 0
        For pX = sx To sx + PixSize - 1
          For pY = sy To sy + PixSize - 1
            If (pX <= PicInfo.bmWidth) And (pY <= PicInfo.bmHeight) Then
              b = b + CLng(bDATA(1, pX, pY))
              g = g + CLng(bDATA(2, pX, pY))
              r = r + CLng(bDATA(3, pX, pY))
              mC = mC + 1
            End If
          Next pY
        Next pX
        If mC > 0 Then
          b = b \ mC
          g = g \ mC
          r = r \ mC
        End If
      End If
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub SwapBank(ByRef pProgress As Long, ByVal Modo As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long

  mProgress = 0
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      Select Case Modo
        Case 1: 'RGB -> BRG
          iDATA(1, X, Y) = g
          iDATA(2, X, Y) = r
          iDATA(3, X, Y) = b
        Case 2: 'RGB -> GBR
          iDATA(1, X, Y) = r
          iDATA(2, X, Y) = b
          iDATA(3, X, Y) = g
        Case 3: 'RGB -> RBG
          iDATA(1, X, Y) = g
          iDATA(2, X, Y) = b
          iDATA(3, X, Y) = r
        Case 4: 'RGB -> BGR
          iDATA(1, X, Y) = r
          iDATA(2, X, Y) = g
          iDATA(3, X, Y) = b
        Case 5: 'RGB -> GRB
          iDATA(1, X, Y) = b
          iDATA(2, X, Y) = r
          iDATA(3, X, Y) = g
      End Select
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub NearestColorBW(ByRef pProgress As Long, ByVal BackCol As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim cB As Long, cG As Long, cR As Long

  Call GetRGB(BackCol, cR, cG, cB)
  mProgress = 0
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      If (r < cR) And (g < cG) And (b < cB) Then
        iDATA(1, X, Y) = 0
        iDATA(2, X, Y) = 0
        iDATA(3, X, Y) = 0
      Else
        iDATA(1, X, Y) = 255
        iDATA(2, X, Y) = 255
        iDATA(3, X, Y) = 255
      End If
    Next X
    If mCancel Then GoTo EndProgress
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub EnhancedDiffusionBW(ByRef pProgress As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim Erro As Long, nCol As Long
  Dim mCol As Long

  mProgress = 0
  For X = 0 To 765
    Speed(X) = X \ 3
  Next X
  mCol = 0
  nCol = 0
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      b = Speed(r + g + b)
      mCol = mCol + b
      nCol = nCol + 1
      If mCancel Then GoTo EndProgress
    Next X
  Next Y
  pProgress = 5
  DoEvents
  mCol = mCol \ nCol
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      If (X > 1) And (Y > 1) And (X < PicInfo.bmWidth) And (Y < PicInfo.bmHeight) Then
        b = CLng(bDATA(1, X - 1, Y - 1)) + CLng(bDATA(1, X - 1, Y)) + CLng(bDATA(1, X - 1, Y + 1)) + CLng(bDATA(1, X, Y - 1)) + _
          CLng(bDATA(1, X, Y + 1)) + CLng(bDATA(1, X + 1, Y - 1)) + CLng(bDATA(1, X + 1, Y)) + CLng(bDATA(1, X + 1, Y + 1))
        g = CLng(bDATA(2, X - 1, Y - 1)) + CLng(bDATA(2, X - 1, Y)) + CLng(bDATA(2, X - 1, Y + 1)) + CLng(bDATA(2, X, Y - 1)) + _
          CLng(bDATA(2, X, Y + 1)) + CLng(bDATA(2, X + 1, Y - 1)) + CLng(bDATA(2, X + 1, Y)) + CLng(bDATA(2, X + 1, Y + 1))
        r = CLng(bDATA(3, X - 1, Y - 1)) + CLng(bDATA(3, X - 1, Y)) + CLng(bDATA(3, X - 1, Y + 1)) + CLng(bDATA(3, X, Y - 1)) + _
          CLng(bDATA(3, X, Y + 1)) + CLng(bDATA(3, X + 1, Y - 1)) + CLng(bDATA(3, X + 1, Y)) + CLng(bDATA(3, X + 1, Y + 1))
        b = (10 * CLng(bDATA(1, X, Y)) - b) \ 2
        g = (10 * CLng(bDATA(2, X, Y)) - g) \ 2
        r = (10 * CLng(bDATA(3, X, Y)) - r) \ 2
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
      Else
        b = CLng(bDATA(1, X, Y))
        g = CLng(bDATA(2, X, Y))
        r = CLng(bDATA(3, X, Y))
      End If
      b = Speed(r + g + b)
      b = b + Erro
      If b < 0 Then b = 0
      If b > 255 Then b = 255
      If b < mCol Then nCol = 0 Else nCol = 255
      Erro = (b - nCol) \ 4
      iDATA(1, X, Y) = nCol
      iDATA(2, X, Y) = nCol
      iDATA(3, X, Y) = nCol
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = 5 + (Y * 95) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub OrderedDitherBW(ByRef pProgress As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim VecDither(1 To 4, 1 To 4) As Byte
  Dim cx As Long, cy As Long

  VecDither(1, 1) = 1:    VecDither(1, 2) = 9
  VecDither(1, 3) = 3:    VecDither(1, 4) = 11
  VecDither(2, 1) = 13:   VecDither(2, 2) = 5
  VecDither(2, 3) = 15:   VecDither(2, 4) = 7
  VecDither(3, 1) = 4:    VecDither(3, 2) = 12
  VecDither(3, 3) = 2:    VecDither(3, 4) = 10
  VecDither(4, 1) = 16:   VecDither(4, 2) = 8
  VecDither(4, 3) = 14:   VecDither(4, 4) = 6
  mProgress = 0
  For X = 0 To 765
    Speed(X) = 1 + (X \ 3) \ 16
  Next X
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      b = Speed(r + g + b)
      cx = 1 + ((X - 1) Mod 4)
      cy = 1 + ((Y - 1) Mod 4)
      If b < VecDither(cx, cy) Then
        iDATA(1, X, Y) = 0
        iDATA(2, X, Y) = 0
        iDATA(3, X, Y) = 0
      Else
        iDATA(1, X, Y) = 255
        iDATA(2, X, Y) = 255
        iDATA(3, X, Y) = 255
      End If
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub FloydSteinbergBW(ByRef pProgress As Long, ByVal PalWeight As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim Erro As Long
  Dim VecErro() As Long
  Dim nCol As Long, mCol As Long
  Dim PartErr(1 To 4, -255 To 255) As Long
  
  For X = 0 To 765
    Speed(X) = X \ 3
  Next X
  For X = -255 To 255
    PartErr(1, X) = (7 * X) \ 16
    PartErr(2, X) = (3 * X) \ 16
    PartErr(3, X) = (5 * X) \ 16
    PartErr(4, X) = (1 * X) \ 16
    If mCancel Then GoTo EndProgress
  Next X
  Erro = 0
  ReDim VecErro(1 To 2, 1 To PicInfo.bmWidth) As Long
  For X = 1 To PicInfo.bmWidth
    VecErro(1, X) = 0
    VecErro(2, X) = 0
    If mCancel Then GoTo EndProgress
  Next X
  pProgress = 2
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      b = Speed(r + g + b)
      mCol = mCol + b
      nCol = nCol + 1
      If mCancel Then GoTo EndProgress
    Next X
  Next Y
  mCol = mCol \ nCol
  pProgress = 10
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      b = Speed(r + g + b)
      b = b + (VecErro(1, X) * 10) \ PalWeight
      If b < 0 Then b = 0
      If b > 255 Then b = 255
      If b < mCol Then nCol = 0 Else nCol = 255
      iDATA(1, X, Y) = nCol
      iDATA(2, X, Y) = nCol
      iDATA(3, X, Y) = nCol
      Erro = b - nCol
      If X < PicInfo.bmWidth Then VecErro(1, X + 1) = VecErro(1, X + 1) + PartErr(1, Erro)
      If Y < PicInfo.bmHeight Then
        If X > 1 Then VecErro(2, X - 1) = VecErro(2, X - 1) + PartErr(2, Erro)
        VecErro(2, X) = VecErro(2, X) + PartErr(3, Erro)
        If X < PicInfo.bmWidth Then VecErro(2, X + 1) = VecErro(2, X + 1) + PartErr(4, Erro)
      End If
      If mCancel Then GoTo EndProgress
    Next X
    For X = 1 To PicInfo.bmWidth
      VecErro(1, X) = VecErro(2, X)
      VecErro(2, X) = 0
    Next X
    mProgress = 10 + (Y * 90) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub BurkeBW(ByRef pProgress As Long, ByVal PalWeight As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim Erro As Long
  Dim VecErro() As Long
  Dim nCol As Long, mCol As Long
  Dim PartErr(1 To 7, -255 To 255) As Long
  
  For X = 0 To 765
    Speed(X) = X \ 3
  Next X
  For X = -255 To 255
    PartErr(1, X) = (8 * X) \ 32
    PartErr(2, X) = (4 * X) \ 32
    PartErr(3, X) = (2 * X) \ 32
    PartErr(4, X) = (4 * X) \ 32
    PartErr(5, X) = (8 * X) \ 32
    PartErr(6, X) = (4 * X) \ 32
    PartErr(7, X) = (2 * X) \ 32
    If mCancel Then GoTo EndProgress
  Next X
  Erro = 0
  ReDim VecErro(1 To 2, 1 To PicInfo.bmWidth) As Long
  For X = 1 To PicInfo.bmWidth
    VecErro(1, X) = 0
    VecErro(2, X) = 0
  Next X
  pProgress = 3
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      b = Speed(r + g + b)
      mCol = mCol + b
      nCol = nCol + 1
      If mCancel Then GoTo EndProgress
    Next X
  Next Y
  mCol = mCol \ nCol
  pProgress = 10
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      b = Speed(r + g + b)
      b = b + (VecErro(1, X) * 10) \ PalWeight
      If b < 0 Then b = 0
      If b > 255 Then b = 255
      If b < mCol Then nCol = 0 Else nCol = 255
      iDATA(1, X, Y) = nCol
      iDATA(2, X, Y) = nCol
      iDATA(3, X, Y) = nCol
      Erro = (b - nCol)
      If X < PicInfo.bmWidth Then VecErro(1, X + 1) = VecErro(1, X + 1) + PartErr(1, Erro)
      If X < PicInfo.bmWidth - 1 Then VecErro(1, X + 2) = VecErro(1, X + 2) + PartErr(2, Erro)
      If Y < PicInfo.bmHeight Then
        If X > 2 Then VecErro(2, X - 2) = VecErro(2, X - 2) + PartErr(3, Erro)
        If X > 1 Then VecErro(2, X - 1) = VecErro(2, X - 1) + PartErr(4, Erro)
        VecErro(2, X) = VecErro(2, X) + PartErr(5, Erro)
        If X < PicInfo.bmWidth Then VecErro(2, X + 1) = VecErro(2, X + 1) + PartErr(6, Erro)
        If X < PicInfo.bmWidth - 1 Then VecErro(2, X + 2) = VecErro(2, X + 2) + PartErr(7, Erro)
      End If
      If mCancel Then GoTo EndProgress
    Next X
    For X = 1 To PicInfo.bmWidth
      VecErro(1, X) = VecErro(2, X)
      VecErro(2, X) = 0
    Next X
    mProgress = 10 + (Y * 90) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub StuckiBW(ByRef pProgress As Long, ByVal PalWeight As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim Erro As Long
  Dim VecErro() As Long
  Dim nCol As Long, mCol As Long
  Dim PartErr(1 To 12, -255 To 255) As Long
  
  For X = 0 To 765
    Speed(X) = X \ 3
  Next X
  For X = -255 To 255
    PartErr(1, X) = (8 * X) \ 42
    PartErr(2, X) = (4 * X) \ 42
    PartErr(3, X) = (2 * X) \ 42
    PartErr(4, X) = (4 * X) \ 42
    PartErr(5, X) = (8 * X) \ 42
    PartErr(6, X) = (4 * X) \ 42
    PartErr(7, X) = (2 * X) \ 42
    PartErr(8, X) = (1 * X) \ 42
    PartErr(9, X) = (2 * X) \ 42
    PartErr(10, X) = (4 * X) \ 42
    PartErr(11, X) = (2 * X) \ 42
    PartErr(12, X) = (1 * X) \ 42
    If mCancel Then GoTo EndProgress
  Next X
  Erro = 0
  ReDim VecErro(1 To 3, 1 To PicInfo.bmWidth) As Long
  For X = 1 To PicInfo.bmWidth
    VecErro(1, X) = 0
    VecErro(2, X) = 0
    VecErro(3, X) = 0
  Next X
  pProgress = 3
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      b = Speed(r + g + b)
      mCol = mCol + b
      nCol = nCol + 1
      If mCancel Then GoTo EndProgress
    Next X
  Next Y
  mCol = mCol \ nCol
  pProgress = 10
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      b = Speed(r + g + b)
      b = b + (VecErro(1, X) * 10) \ PalWeight
      If b < 0 Then b = 0
      If b > 255 Then b = 255
      If b < mCol Then nCol = 0 Else nCol = 255
      iDATA(1, X, Y) = nCol
      iDATA(2, X, Y) = nCol
      iDATA(3, X, Y) = nCol
      Erro = (b - nCol)
      If X < PicInfo.bmWidth Then VecErro(1, X + 1) = VecErro(1, X + 1) + PartErr(1, Erro)
      If X < PicInfo.bmWidth - 1 Then VecErro(1, X + 2) = VecErro(1, X + 2) + PartErr(2, Erro)
      If Y < PicInfo.bmHeight Then
        If X > 2 Then VecErro(2, X - 2) = VecErro(2, X - 2) + PartErr(3, Erro)
        If X > 1 Then VecErro(2, X - 1) = VecErro(2, X - 1) + PartErr(4, Erro)
        VecErro(2, X) = VecErro(2, X) + PartErr(5, Erro)
        If X < PicInfo.bmWidth Then VecErro(2, X + 1) = VecErro(2, X + 1) + PartErr(6, Erro)
        If X < PicInfo.bmWidth - 1 Then VecErro(2, X + 2) = VecErro(2, X + 2) + PartErr(7, Erro)
      End If
      If Y < PicInfo.bmHeight - 1 Then
        If X > 2 Then VecErro(3, X - 2) = VecErro(3, X - 2) + PartErr(8, Erro)
        If X > 1 Then VecErro(3, X - 1) = VecErro(3, X - 1) + PartErr(9, Erro)
        VecErro(3, X) = VecErro(3, X) + PartErr(10, Erro)
        If X < PicInfo.bmWidth Then VecErro(3, X + 1) = VecErro(3, X + 1) + PartErr(11, Erro)
        If X < PicInfo.bmWidth - 1 Then VecErro(3, X + 2) = VecErro(3, X + 2) + PartErr(12, Erro)
      End If
      If mCancel Then GoTo EndProgress
    Next X
    For X = 1 To PicInfo.bmWidth
      VecErro(1, X) = VecErro(2, X)
      VecErro(2, X) = VecErro(3, X)
      VecErro(3, X) = 0
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = 10 + (Y * 90) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Aqua(ByRef pProgress As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim Med(1 To 4) As Long
  Dim Dev(1 To 4) As Long
  Dim i As Long, j As Long
  Dim sDev As Long, vDev As Long
  
  mProgress = 0
  For Y = 3 To PicInfo.bmHeight - 2
    For X = 3 To PicInfo.bmWidth - 2
      For i = 1 To 3
        Med(1) = CLng(bDATA(i, X - 2, Y - 2)) + CLng(bDATA(i, X - 1, Y - 2)) + CLng(bDATA(i, X, Y - 2)) + _
                 CLng(bDATA(i, X - 2, Y - 1)) + CLng(bDATA(i, X - 1, Y - 1)) + CLng(bDATA(i, X, Y - 1)) + _
                 CLng(bDATA(i, X - 2, Y)) + CLng(bDATA(i, X - 1, Y)) + CLng(bDATA(i, X, Y))
        Med(2) = CLng(bDATA(i, X + 2, Y - 2)) + CLng(bDATA(i, X + 1, Y - 2)) + CLng(bDATA(i, X, Y - 2)) + _
                 CLng(bDATA(i, X + 2, Y - 1)) + CLng(bDATA(i, X + 1, Y - 1)) + CLng(bDATA(i, X, Y - 1)) + _
                 CLng(bDATA(i, X + 2, Y)) + CLng(bDATA(i, X + 1, Y)) + CLng(bDATA(i, X, Y))
        Med(3) = CLng(bDATA(i, X - 2, Y + 2)) + CLng(bDATA(i, X - 1, Y + 2)) + CLng(bDATA(i, X, Y + 2)) + _
                 CLng(bDATA(i, X - 2, Y + 1)) + CLng(bDATA(i, X - 1, Y + 1)) + CLng(bDATA(i, X, Y + 1)) + _
                 CLng(bDATA(i, X - 2, Y)) + CLng(bDATA(i, X - 1, Y)) + CLng(bDATA(i, X, Y))
        Med(4) = CLng(bDATA(i, X + 2, Y + 2)) + CLng(bDATA(i, X + 1, Y + 2)) + CLng(bDATA(i, X, Y + 2)) + _
                 CLng(bDATA(i, X + 2, Y + 1)) + CLng(bDATA(i, X + 1, Y + 1)) + CLng(bDATA(i, X, Y + 1)) + _
                 CLng(bDATA(i, X + 2, Y)) + CLng(bDATA(i, X + 1, Y)) + CLng(bDATA(i, X, Y))
        Med(1) = Med(1) \ 9
        Med(2) = Med(2) \ 9
        Med(3) = Med(3) \ 9
        Med(4) = Med(4) \ 9
        Dev(1) = Abs(CLng(bDATA(i, X - 2, Y - 2)) - Med(1)) + Abs(CLng(bDATA(i, X - 1, Y - 2)) - Med(1)) + Abs(CLng(bDATA(i, X, Y - 2)) - Med(1)) + _
                 Abs(CLng(bDATA(i, X - 2, Y - 1)) - Med(1)) + Abs(CLng(bDATA(i, X - 1, Y - 1)) - Med(1)) + Abs(CLng(bDATA(i, X, Y - 1)) - Med(1)) + _
                 Abs(CLng(bDATA(i, X - 2, Y)) - Med(1)) + Abs(CLng(bDATA(i, X - 1, Y)) - Med(1)) + Abs(CLng(bDATA(i, X, Y)) - Med(1))
        Dev(2) = Abs(CLng(bDATA(i, X + 2, Y - 2)) - Med(2)) + Abs(CLng(bDATA(i, X + 1, Y - 2)) - Med(2)) + Abs(CLng(bDATA(i, X, Y - 2)) - Med(2)) + _
                 Abs(CLng(bDATA(i, X + 2, Y - 1)) - Med(2)) + Abs(CLng(bDATA(i, X + 1, Y - 1)) - Med(2)) + Abs(CLng(bDATA(i, X, Y - 1)) - Med(2)) + _
                 Abs(CLng(bDATA(i, X + 2, Y)) - Med(2)) + Abs(CLng(bDATA(i, X + 1, Y)) - Med(2)) + Abs(CLng(bDATA(i, X, Y)) - Med(2))
        Dev(3) = Abs(CLng(bDATA(i, X - 2, Y + 2)) - Med(3)) + Abs(CLng(bDATA(i, X - 1, Y + 2)) - Med(3)) + Abs(CLng(bDATA(i, X, Y + 2)) - Med(3)) + _
                 Abs(CLng(bDATA(i, X - 2, Y + 1)) - Med(3)) + Abs(CLng(bDATA(i, X - 1, Y + 1)) - Med(3)) + Abs(CLng(bDATA(i, X, Y + 1)) - Med(3)) + _
                 Abs(CLng(bDATA(i, X - 2, Y)) - Med(3)) + Abs(CLng(bDATA(i, X - 1, Y)) - Med(3)) + Abs(CLng(bDATA(i, X, Y)) - Med(3))
        Dev(4) = Abs(CLng(bDATA(i, X + 2, Y + 2)) - Med(4)) + Abs(CLng(bDATA(i, X + 1, Y + 2)) - Med(4)) + Abs(CLng(bDATA(i, X, Y + 2)) - Med(4)) + _
                 Abs(CLng(bDATA(i, X + 2, Y + 1)) - Med(4)) + Abs(CLng(bDATA(i, X + 1, Y + 1)) - Med(4)) + Abs(CLng(bDATA(i, X, Y + 1)) - Med(4)) + _
                 Abs(CLng(bDATA(i, X + 2, Y)) - Med(4)) + Abs(CLng(bDATA(i, X + 1, Y)) - Med(4)) + Abs(CLng(bDATA(i, X, Y)) - Med(4))
        vDev = 99999
        sDev = 0
        For j = 1 To 4
          If Dev(j) < vDev Then
            vDev = Dev(j)
            sDev = j
          End If
        Next j
        iDATA(i, X, Y) = Med(sDev)
        If mCancel Then GoTo EndProgress
      Next i
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Dilate(ByRef pProgress As Long)
  Dim X As Long, Y As Long
  Dim v As Long
  Dim i As Long
  Dim vMax As Long
  On Error Resume Next
  mProgress = 0
  For Y = 1 To PicInfo.bmHeight '- 1
    For X = 1 To PicInfo.bmWidth '- 1
      For i = 1 To 3
        vMax = 0
        v = CLng(bDATA(i, X - 1, Y - 1))
        If v > vMax Then vMax = v
        v = CLng(bDATA(i, X, Y - 1))
        If v > vMax Then vMax = v
        v = CLng(bDATA(i, X + 1, Y - 1))
        If v > vMax Then vMax = v
        
        v = CLng(bDATA(i, X - 1, Y))
        If v > vMax Then vMax = v
        v = CLng(bDATA(i, X, Y))
        If v > vMax Then vMax = v
        v = CLng(bDATA(i, X + 1, Y))
        If v > vMax Then vMax = v
        
        v = CLng(bDATA(i, X - 1, Y + 1))
        If v > vMax Then vMax = v
        v = CLng(bDATA(i, X, Y + 1))
        If v > vMax Then vMax = v
        v = CLng(bDATA(i, X + 1, Y + 1))
        If v > vMax Then vMax = v
        
        iDATA(i, X, Y) = vMax
      Next i
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Erode(ByRef pProgress As Long)
  Dim X As Long, Y As Long
  Dim v As Long
  Dim i As Long
  Dim vMin As Long
  On Error Resume Next
  mProgress = 0
  For Y = 2 To PicInfo.bmHeight - 1
    For X = 2 To PicInfo.bmWidth - 1
      For i = 1 To 3
        vMin = 255
        v = CLng(bDATA(i, X - 1, Y - 1))
        If v < vMin Then vMin = v
        v = CLng(bDATA(i, X, Y - 1))
        If v < vMin Then vMin = v
        v = CLng(bDATA(i, X + 1, Y - 1))
        If v < vMin Then vMin = v
        
        v = CLng(bDATA(i, X - 1, Y))
        If v < vMin Then vMin = v
        v = CLng(bDATA(i, X, Y))
        If v < vMin Then vMin = v
        v = CLng(bDATA(i, X + 1, Y))
        If v < vMin Then vMin = v
        
        v = CLng(bDATA(i, X - 1, Y + 1))
        If v < vMin Then vMin = v
        v = CLng(bDATA(i, X, Y + 1))
        If v < vMin Then vMin = v
        v = CLng(bDATA(i, X + 1, Y + 1))
        If v < vMin Then vMin = v
        
        iDATA(i, X, Y) = vMin
      Next i
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub ConnectedContour(ByRef pProgress As Long)
  Dim X As Long, Y As Long
  Dim v As Long
  Dim i As Long
  Dim vMin As Long
   On Error Resume Next
  mProgress = 0
  For Y = 1 To PicInfo.bmHeight '- 1
    For X = 1 To PicInfo.bmWidth '- 1
      For i = 1 To 3
        vMin = 255
        v = CLng(bDATA(i, X - 1, Y - 1))
        If v < vMin Then vMin = v
        v = CLng(bDATA(i, X, Y - 1))
        If v < vMin Then vMin = v
        v = CLng(bDATA(i, X + 1, Y - 1))
        If v < vMin Then vMin = v
        
        v = CLng(bDATA(i, X - 1, Y))
        If v < vMin Then vMin = v
        v = CLng(bDATA(i, X, Y))
        If v < vMin Then vMin = v
        v = CLng(bDATA(i, X + 1, Y))
        If v < vMin Then vMin = v
        
        v = CLng(bDATA(i, X - 1, Y + 1))
        If v < vMin Then vMin = v
        v = CLng(bDATA(i, X, Y + 1))
        If v < vMin Then vMin = v
        v = CLng(bDATA(i, X + 1, Y + 1))
        If v < vMin Then vMin = v
        
        iDATA(i, X, Y) = CLng(iDATA(i, X, Y)) - vMin
      Next i
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub StretchHistogram(ByRef pProgress As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim bMin As Long, bMax As Long
  Dim gMin As Long, gMax As Long
  Dim rMin As Long, rMax As Long
  
  mProgress = 0
  bMin = 255: bMax = 0
  gMin = 255: gMax = 0
  rMin = 255: rMax = 0
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      If b < bMin Then bMin = b
      If b > bMax Then bMax = b
      If g < gMin Then gMin = g
      If g > gMax Then gMax = g
      If r < rMin Then rMin = r
      If r > rMax Then rMax = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 10) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 10
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      b = 255 * (b - bMin) / (bMax - bMin)
      g = 255 * (g - gMin) / (gMax - gMin)
      r = 255 * (r - rMin) / (rMax - rMin)
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = 10 + (Y * 90) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub AddNoise(ByRef pProgress As Long, ByVal Factor As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim v As Long
    
  mProgress = 0
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y)) + ((Factor * 2 + 1) * Rnd - Factor)
      g = CLng(bDATA(2, X, Y)) + ((Factor * 2 + 1) * Rnd - Factor)
      r = CLng(bDATA(3, X, Y)) + ((Factor * 2 + 1) * Rnd - Factor)
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If g > 255 Then g = 255
      If g < 0 Then g = 0
      If b > 255 Then b = 255
      If b < 0 Then b = 0
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Saturation(ByRef pProgress As Long, ByVal Factor As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim v As Long
  Dim sF As Single
    
  mProgress = 0
  For X = 0 To 765
    Speed(X) = X \ 3
  Next X
  sF = Factor / 100
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      v = Speed(b + g + r)
      b = b + sF * (b - v)
      g = g + sF * (g - v)
      r = r + sF * (r - v)
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If g > 255 Then g = 255
      If g < 0 Then g = 0
      If b > 255 Then b = 255
      If b < 0 Then b = 0
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub GammaCorrection(ByRef pProgress As Long, ByVal Factor As Long)
  Dim X As Long, Y As Long
  Dim r As Long, g As Long, b As Long
  Dim dB As Double, dG As Double, dR As Double
  Dim sF As Single
  Dim Max As Double, Min As Double, MM As Double
  Dim H As Double, s As Double, i As Double
  Dim cB As Double, cG As Double, cR As Double
  Dim Flo As Long
    
  mProgress = 0
  sF = Factor / 100
  For Y = 1 To PicInfo.bmHeight
    For X = 1 To PicInfo.bmWidth
      'get data
      b = CLng(bDATA(1, X, Y))
      g = CLng(bDATA(2, X, Y))
      r = CLng(bDATA(3, X, Y))
      dB = b / 255
      dG = g / 255
      dR = r / 255
      'correct gamma
      dB = dB ^ (1 / sF)
      dG = dG ^ (1 / sF)
      dR = dR ^ (1 / sF)
      'set data
      b = dB * 255
      g = dG * 255
      r = dR * 255
      iDATA(1, X, Y) = b
      iDATA(2, X, Y) = g
      iDATA(3, X, Y) = r
      If mCancel Then GoTo EndProgress
    Next X
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
EndProgress:
  pProgress = 100
  DoEvents
End Sub

Private Sub Neon(ByRef pProgress As Long)

     Dim VPic(3) As Integer
     Dim i As Long, j As Long, k As Integer
     Dim Color_2 As Long, Color_1 As Long
     Dim b As Long, g As Long
     On Error Resume Next
     pProgress = 0
     
     For i = 1 To PicInfo.bmHeight '- 1
        For j = 1 To PicInfo.bmWidth '- 1
            For k = 1 To 3
                b = CLng(bDATA(k, j, i))
                g = CLng(bDATA(k, j + 1, i))
                Color_1 = (b - g) ^ 2
                g = CLng(bDATA(k, j, i + 1))
                Color_2 = (b - g) ^ 2
                VPic(k) = 2 * (Color_1 + Color_2) ^ 0.5
                If VPic(k) > 255 Then VPic(k) = 255
                If VPic(k) < 0 Then VPic(k) = 0
                iDATA(k, j, i) = VPic(k)
            Next
            If mCancel Then GoTo EndProgress
            mProgress = (i * 100) \ PicInfo.bmHeight
            pProgress = mProgress
            DoEvents
        Next
    Next
EndProgress:
    pProgress = 100
    
End Sub

'3d Grid
Private Sub Grid3d(ByRef pProgress As Long, ByVal Factor As Long)
  Dim X As Long, Y As Long, Counter As Long
  Dim r As Long, g As Long, b As Long
  On Error Resume Next
  pProgress = 0
  For Y = 1 To PicInfo.bmHeight Step Factor + 1
    For X = 1 To PicInfo.bmWidth Step Factor + 1
        b = CLng(bDATA(1, X, Y))
        g = CLng(bDATA(2, X, Y))
        r = CLng(bDATA(3, X, Y))
        r = r - 20 '
        g = g - 20 '
        b = b - 20 '
        For Counter = 1 To Factor
            iDATA(1, X + Counter, Y) = b
            iDATA(2, X + Counter, Y) = g
            iDATA(3, X + Counter, Y) = r
            
            iDATA(1, X - Counter, Y) = b
            iDATA(2, X - Counter, Y) = g
            iDATA(3, X - Counter, Y) = r
            
            iDATA(1, X, Y + Counter) = b
            iDATA(2, X, Y + Counter) = g
            iDATA(3, X, Y + Counter) = r
            
            iDATA(1, X, Y + Counter) = b
            iDATA(2, X, Y + Counter) = g
            iDATA(3, X, Y + Counter) = r
        Next
        If mCancel Then GoTo EndProgress
        mProgress = (Y * 100) \ PicInfo.bmHeight
        pProgress = mProgress
        DoEvents
    Next
EndProgress:
    pProgress = 100
 Next
 
End Sub

'mirror left right
Private Sub MirrorLeftRight(ByRef pProgress As Long)
   Dim X As Long, Y As Long
   Dim r As Long, g As Long, b As Long
 
   On Error Resume Next
    pProgress = 0
    For Y = 1 To PicInfo.bmHeight
        For X = 1 To PicInfo.bmWidth / 2
         b = CLng(bDATA(1, X, Y))
         g = CLng(bDATA(2, X, Y))
         r = CLng(bDATA(3, X, Y))
         
         iDATA(1, PicInfo.bmWidth - X, Y) = b
         iDATA(2, PicInfo.bmWidth - X, Y) = g
         iDATA(3, PicInfo.bmWidth - X, Y) = r
         If mCancel Then GoTo EndProgress
      Next
      mProgress = (Y * 100) \ PicInfo.bmHeight
      pProgress = mProgress
      DoEvents
    Next
EndProgress:
    pProgress = 100

End Sub

'mirror right left
Private Sub MirrorRightLeft(ByRef pProgress As Long)
   Dim X As Long, Y As Long
   Dim r As Long, g As Long, b As Long
    pProgress = 0
   On Error Resume Next
    For Y = 1 To PicInfo.bmHeight
        For X = 1 To PicInfo.bmWidth / 2
         b = CLng(bDATA(1, PicInfo.bmWidth - X, Y))
         g = CLng(bDATA(2, PicInfo.bmWidth - X, Y))
         r = CLng(bDATA(3, PicInfo.bmWidth - X, Y))
         iDATA(1, X, Y) = b
         iDATA(2, X, Y) = g
         iDATA(3, X, Y) = r
         If mCancel Then GoTo EndProgress
      Next
        mProgress = (Y * 100) \ PicInfo.bmHeight
        pProgress = mProgress
        DoEvents
    Next
EndProgress:
    pProgress = 100

End Sub

'mirror top down
Private Sub MirrorTopDown(ByRef pProgress As Long)
   Dim X As Long, Y As Long
   Dim r As Long, g As Long, b As Long
    pProgress = 0
   On Error Resume Next
    For Y = 1 To PicInfo.bmHeight / 2
        For X = 1 To PicInfo.bmWidth
         b = CLng(bDATA(1, X, Y))
         g = CLng(bDATA(2, X, Y))
         r = CLng(bDATA(3, X, Y))
         iDATA(1, X, PicInfo.bmHeight - Y) = b
         iDATA(2, X, PicInfo.bmHeight - Y) = g
         iDATA(3, X, PicInfo.bmHeight - Y) = r
         If mCancel Then GoTo EndProgress
      Next
        mProgress = (Y * 100) \ PicInfo.bmHeight
        pProgress = mProgress
        DoEvents
    Next
EndProgress:
    pProgress = 100

End Sub

'mirror down top
Private Sub MirrorDownTop(ByRef pProgress As Long)
   Dim X As Long, Y As Long
   Dim r As Long, g As Long, b As Long
   pProgress = 0
   On Error Resume Next
    For Y = 1 To PicInfo.bmHeight / 2
        For X = 1 To PicInfo.bmWidth
         b = CLng(bDATA(1, X, PicInfo.bmHeight - Y))
         g = CLng(bDATA(2, X, PicInfo.bmHeight - Y))
         r = CLng(bDATA(3, X, PicInfo.bmHeight - Y))
         iDATA(1, X, Y) = b
         iDATA(2, X, Y) = g
         iDATA(3, X, Y) = r
         If mCancel Then GoTo EndProgress
      Next
      mProgress = (Y * 100) \ PicInfo.bmHeight
      pProgress = mProgress
      DoEvents
    Next
EndProgress:
    pProgress = 100

End Sub

'Art
Private Sub Art(ByRef pProgress As Long)
   Dim X As Long, Y As Long
   Dim r As Long, g As Long, b As Long
   Dim Rr1 As Long, Gg1 As Long, Bb1 As Long
   On Error Resume Next
    pProgress = 0
   For Y = 1 To PicInfo.bmHeight
      For X = 1 To PicInfo.bmWidth
         b = CLng(bDATA(1, X, Y))
         g = CLng(bDATA(2, X, Y))
         r = CLng(bDATA(3, X, Y))
         
        If r > 127 Then
            Rr1 = 255 - r
            r = r / Rr1
        Else
            Rr1 = 255 - r
            r = r * Rr1
        End If

        If g > 127 Then
            Gg1 = 255 - g
            g = g / Gg1
        Else
            Gg1 = 255 - g
            g = g * Gg1
        End If

        If b > 127 Then
            Bb1 = 255 - b
            b = b / Bb1
        Else
            Bb1 = 255 - b
            b = b * Bb1
        End If
        iDATA(1, X, Y) = b
        iDATA(2, X, Y) = g
        iDATA(3, X, Y) = r
        If mCancel Then GoTo EndProgress
    Next
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next
EndProgress:
  pProgress = 100
End Sub

'Stranges
Private Sub Stranges(ByRef pProgress As Long, ByVal Factor As Long)
   Dim X As Long, Y As Long
   Dim r As Long, g As Long, b As Long
   Dim Rr1 As Long, Gg1 As Long, Bb1 As Long
   On Error Resume Next
    pProgress = 0
   For Y = 1 To PicInfo.bmHeight
      For X = 1 To PicInfo.bmWidth
         b = CLng(bDATA(1, X, Y))
         g = CLng(bDATA(2, X, Y))
         r = CLng(bDATA(3, X, Y))
         
         If r > 127 Then
            r = 255 - r / Factor
         Else
            r = 0 + r / Factor
         End If

         If g > 127 Then
            g = 255 - g / Factor
         Else
            g = 0 + g / Factor
         End If

         If b > 127 Then
            b = 255 - b / Factor
         Else
            b = 0 + b / Factor
         End If

        iDATA(1, X, Y) = b
        iDATA(2, X, Y) = g
        iDATA(3, X, Y) = r
        If mCancel Then GoTo EndProgress
    Next
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next
EndProgress:
    pProgress = 100
End Sub

'Fog
Private Sub Fog(ByRef pProgress As Long, ByVal Factor As Long)

   Dim X As Long, Y As Long
   Dim r As Long, g As Long, b As Long
   Dim Rr1 As Long, Gg1 As Long, Bb1 As Long
   On Error Resume Next
    pProgress = 0
   For Y = 1 To PicInfo.bmHeight
      For X = 1 To PicInfo.bmWidth
         b = CLng(bDATA(1, X, Y))
         g = CLng(bDATA(2, X, Y))
         r = CLng(bDATA(3, X, Y))
   
        If Val(r) > 127 Then
            r = r - Factor
            If r < 127 Then r = 127
        Else
            r = r + Factor
            If r > 127 Then r = 127
        End If
   
        If Val(g) > 127 Then
            g = g - Factor
            If g < 127 Then g = 127
        Else
            g = g + Factor
            If g > 127 Then g = 127
        End If
    
        If Val(b) > 127 Then
            b = b - Factor
            If b < 127 Then b = 127
        Else
            b = b + Factor
            If b > 127 Then b = 127
        End If
    
        iDATA(1, X, Y) = b
        iDATA(2, X, Y) = g
        iDATA(3, X, Y) = r
        If mCancel Then GoTo EndProgress
    Next
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next
EndProgress:
  pProgress = 100
End Sub

'Snow 4-64
Private Sub Snow(ByRef pProgress As Long, ByVal Factor As Long)
 
   Dim X As Long, Y As Long, lngWriteColor As Long
   Dim r As Long, g As Long, b As Long
   Dim Rr1 As Long, Gg1 As Long, Bb1 As Long
   On Error Resume Next
    pProgress = 0
   For Y = 1 To PicInfo.bmHeight '- 1
      For X = 1 To PicInfo.bmWidth '- 1
           b = bDATA(1, X, Y) + (Factor * (bDATA(1, X, Y) - bDATA(1, X - 1, Y - 1)))
           g = bDATA(2, X, Y) + (Factor * (bDATA(2, X, Y) - bDATA(2, X - 1, Y - 1)))
           r = bDATA(3, X, Y) + (Factor * (bDATA(3, X, Y) - bDATA(3, X - 1, Y - 1)))
           lngWriteColor = RGB(Abs(r), Abs(g), Abs(b))
           GetRGB lngWriteColor, r, g, b
           iDATA(1, X, Y) = (b)
           iDATA(2, X, Y) = (g)
           iDATA(3, X, Y) = (r)
           If mCancel Then GoTo EndProgress
      Next
      mProgress = (Y * 100) \ PicInfo.bmHeight
      pProgress = mProgress
      DoEvents
   Next
EndProgress:
   pProgress = 100
End Sub

'Wave 0-16
Private Sub Wave(ByRef pProgress As Long, ByVal Factor As Long)
      
   Dim X As Long, Y As Long
   Dim r As Long, g As Long, b As Long
   
   On Error Resume Next
    pProgress = 0
   For Y = 1 To PicInfo.bmHeight
      For X = 1 To PicInfo.bmWidth
          b = bDATA(1, X, Y)
          g = bDATA(2, X, Y)
          r = bDATA(3, X, Y)
          If (Sin(X) * Factor) + Y > 0 Then
            iDATA(1, X, (Sin(X) * Factor) + Y) = (b)
            iDATA(2, X, (Sin(X) * Factor) + Y) = (g)
            iDATA(3, X, (Sin(X) * Factor) + Y) = (r)
          End If
          If mCancel Then GoTo EndProgress
      Next
      mProgress = (Y * 100) \ PicInfo.bmHeight
      pProgress = mProgress
      DoEvents
   Next
EndProgress:
   pProgress = 100
   
End Sub

'sepia
Private Sub Sepia(ByRef pProgress As Long)
   Dim X As Long, Y As Long, lngWriteColor As Long
   Dim r As Long, g As Long, b As Long
   
   On Error Resume Next
   pProgress = 0
   For Y = 1 To PicInfo.bmHeight
      For X = 1 To PicInfo.bmWidth
          b = bDATA(1, X, Y) * 0.114
          g = bDATA(2, X, Y) * 0.587
          r = bDATA(3, X, Y) * 0.299
          lngWriteColor = b + g + r
          b = lngWriteColor
          g = lngWriteColor
          r = lngWriteColor
         If r < 63 Then r = r * 1.1: b = b * 0.9
         If r > 62 And r < 192 Then r = r * 1.15: b = b * 0.85
         If r > 191 Then
            r = r * 1.08
            If r > 255 Then r = 255
            b = b * 0.93
         End If
         iDATA(1, X, Y) = b
         iDATA(3, X, Y) = r
         iDATA(2, X, Y) = g
         If mCancel Then GoTo EndProgress
       Next
      mProgress = (Y * 100) \ PicInfo.bmHeight
      pProgress = mProgress
      DoEvents
   Next
EndProgress:
   pProgress = 100
      
End Sub

Private Sub Rects(ByRef pProgress As Long)
  Dim X As Long, Y As Long, NewColor As Long
  Dim r As Long, g As Long, b As Long
  Dim tR As Long, tG As Long, tB As Long
  Dim tC1 As Long, tC2 As Long, tC3 As Long, tC4 As Long, tC5 As Long
   On Error Resume Next
   pProgress = 0
   For Y = 1 To PicInfo.bmHeight
      For X = 1 To PicInfo.bmWidth
          b = bDATA(1, X, Y)
          g = bDATA(2, X, Y)
          r = bDATA(3, X, Y)
           tC1 = RGB(r, g, b)
          If X <> PicInfo.bmWidth Then
            b = bDATA(1, X + 1, Y)
            g = bDATA(2, X + 1, Y)
            r = bDATA(3, X + 1, Y)
          End If
          tC2 = RGB(r, g, b)
          If X > 1 Then
            b = bDATA(1, X - 1, Y)
            g = bDATA(2, X - 1, Y)
            r = bDATA(3, X - 1, Y)
          End If
          tC3 = RGB(r, g, b)
          If Y <> PicInfo.bmHeight Then
            b = bDATA(1, X, Y + 1)
            g = bDATA(2, X, Y + 1)
            r = bDATA(3, X, Y + 1)
          End If
          tC4 = RGB(r, g, b)
          If Y > 1 Then
          b = bDATA(1, X, Y - 1)
          g = bDATA(2, X, Y - 1)
          r = bDATA(3, X, Y - 1)
          End If
          tC5 = RGB(r, g, b)
          NewColor = Abs(tC1) - (Abs(tC2 + tC3 + tC4 + tC5) / 4)
          GetRGB NewColor, r, g, b
          If r < 0 Then r = 0
          If g < 0 Then g = 0
          If b < 0 Then b = 0
          If r > 255 Then r = 255
          If g > 255 Then g = 255
          If b > 255 Then b = 255
         iDATA(1, X, Y) = b
         iDATA(2, X, Y) = r
         iDATA(3, X, Y) = g
         If mCancel Then GoTo EndProgress
       Next
      mProgress = (Y * 100) \ PicInfo.bmHeight
      pProgress = mProgress
      DoEvents
   Next
EndProgress:
   pProgress = 100
End Sub

'Ice
Private Sub Ice(ByRef pProgress As Long)
 
   Dim X As Long, Y As Long, lngWriteColor As Long
   Dim r As Long, g As Long, b As Long
   Dim Rr1 As Long, Gg1 As Long, Bb1 As Long
   On Error Resume Next
    pProgress = 0
   For Y = 1 To PicInfo.bmHeight
      For X = 1 To PicInfo.bmWidth
           b = bDATA(1, X, Y)
           g = bDATA(2, X, Y)
           r = bDATA(3, X, Y)

           r = Abs((r - g - b) * 1.5)
           g = Abs((g - b - r) * 1.5)
           b = Abs((b - r - g) * 1.5)
           If r > 255 Then r = 255
           If r < 0 Then r = 0
           If g > 255 Then g = 255
           If g < 0 Then g = 0
           If b > 255 Then b = 255
           If b < 0 Then b = 0
           iDATA(1, X, Y) = b
           iDATA(2, X, Y) = g
           iDATA(3, X, Y) = r
           If mCancel Then GoTo EndProgress
      Next
      mProgress = (Y * 100) \ PicInfo.bmHeight
      pProgress = mProgress
      DoEvents
   Next
EndProgress:
   pProgress = 100
End Sub

Private Sub Comic(ByRef pProgress As Long)
  Dim X As Long, Y As Long, NewColor As Long
  Dim r As Long, g As Long, b As Long
  
   On Error Resume Next
   pProgress = 0
   For Y = 1 To PicInfo.bmHeight
      For X = 1 To PicInfo.bmWidth
      b = bDATA(1, X, Y)
          g = bDATA(2, X, Y)
          r = bDATA(3, X, Y)
            
          r = Abs(r * (g - b + g + r)) / 256
          g = Abs(r * (b - g + b + r)) / 256
          b = Abs(g * (b - g + b + r)) / 256
            
         NewColor = RGB(r, g, b)
          GetRGB NewColor, r, g, b
            
          r = (r + g + b) / 3
         
          iDATA(1, X, Y) = r
          iDATA(2, X, Y) = r
          iDATA(3, X, Y) = r
         If mCancel Then GoTo EndProgress
       Next
      mProgress = (Y * 100) \ PicInfo.bmHeight
      pProgress = mProgress
      DoEvents
   Next
EndProgress:
   pProgress = 100
End Sub

