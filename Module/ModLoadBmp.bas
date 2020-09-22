Attribute VB_Name = "ModLoadBitmap"
'http://www.vbarchiv.net/archiv/tipp_1268.html
'==============================================================
' Start source modLoadPicBox
' ==============================================================

Option Explicit

'Image file by 'GDIPlus.DLL' in 'VB.PictureBox' load

' ===============================================================
' Required GDIPlus declarations for loading / Draw a picture
' ===============================================================

Private Type GDIPlusStartupInput
  Version As Long
  etctetera(1 To 12) As Byte
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef GDIP_Connection As Long, ByRef udtInput As GDIPlusStartupInput, Optional ByRef udtOutput As Any) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef Graphics As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal Filename As Long, ByRef Image As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long

'Public Function LoadPicBox(ByVal FileName As String, _
                           ByRef picBox As PictureBox, _
                           Optional ByVal AdjustPicBoxSize As Boolean = True, _
                           Optional ByRef Announcement As String, _
                           Optional retcode As Long) As Boolean
    
Public Function LoadPicBox(ByVal Filename As String, _
                           ByRef picBoxhdc As PictureBox, _
                           Optional ByRef Announcement As String, _
                           Optional Bitmap As Long) As Boolean
      
'Opens image file 'Filename' and bears
'Using the device context of Picturebox
'In the canvas of a PictureBox
'
'If AdjustPicBoxSize = True
'Picturebox which will possibly enlarged,
'At the whole picture to draw
    
  Dim retcode As Long              ' Function returns
'  Dim Bitmap As Long
  Dim Graphics As Long
  Dim picWidth As Single           ' Image Dimensions
  Dim picHeight As Single
  Dim GDIP_Connection As Long      ' Connect to GDIPlus
  Dim GDIP_Startup As GDIPlusStartupInput
  Dim w As Long, h As Long         ' Image size in twips
  
  On Error GoTo exitfunction
  Err.Clear
    
'  If FileExists(FileName) = False Then
'  'If Trim(FileName) = "" Or picBox Is Nothing Then
'    Announcement = "Input parameters are missing"
'    Exit Function
'  End If
  If FileExists(Filename) = False Then
  'If Dir(FileName, vbNormal Or vbReadOnly) = "" Then
    Announcement = "File does not exist"
    Exit Function
  End If
     
  Announcement = ""
  GDIP_Startup.Version = 1
  retcode = GdiplusStartup(GDIP_Connection, GDIP_Startup, ByVal 0&)
  If retcode <> 0 Then
     Announcement = "GDIPlus Unavailable"
     Exit Function
  End If
    
  'Does the image from the file into a bitmap
  retcode = GdipLoadImageFromFile(StrPtr(Filename), Bitmap)
  If retcode <> 0 Then
     Announcement = "Bitmap can not be opened"
     GoTo exitfunction
  End If
        
  'Query dimensions of the bitmap
  retcode = GdipGetImageDimension(Bitmap, picWidth, picHeight)
  If retcode <> 0 Then
    Announcement = "Bitmap-Dimensions unavailable"
    GoTo exitfunction
  End If
      
   'PictureBox, and if necessary
   'On the required image dimensions set
'  With picBox
'     .AutoRedraw = True
'     .BorderStyle = 0
'     .Picture = LoadPicture() 'PictureBox clear
'    If AdjustPicBoxSize Then
'       w = .ScaleX(picWidth, vbPixels, vbTwips)
'        h = .ScaleY(picHeight, vbPixels, vbTwips)
'       If .Width < w Then .Width = w
'       If .Height < h Then .Height = h
'    End If
'  End With
        
  'Create a graphic object GDIPlus
  'For use with Hdc
  retcode = GdipCreateFromHDC(picBoxhdc, Graphics)
  If retcode <> 0 Then
     Announcement = "Graphics object Unavailable"
     GoTo exitfunction
  End If
    
  ' Bitmap to draw PictureBox
  retcode = GdipDrawImageRect(Graphics, Bitmap, 0, 0, picWidth, picHeight)
  If retcode <> 0 Then
    Announcement = "Image can not be drawn"
    GoTo exitfunction
  End If
 
  'Return: everything's OK
  Announcement = ""
  LoadPicBox = True
      retcode = Graphics

exitfunction:
  ' error
  If Err.Number <> 0 Then
     Announcement = Err.Description
  End If
    
  ' Resources and GDIPLus Release
  If Bitmap <> 0 Then
    ' Bitmap löschen
    GdipDisposeImage Bitmap
  End If
    
  If Graphics <> 0 Then
    ' GDIPLus-Delete graphic object
    GdipDeleteGraphics Graphics
  End If
    
  If GDIP_Connection <> 0 Then
    ' GDIPlus-DLL Release
    GdiplusShutdown GDIP_Connection
  End If
End Function


