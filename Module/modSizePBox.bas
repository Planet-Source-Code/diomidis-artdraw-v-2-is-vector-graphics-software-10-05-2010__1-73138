Attribute VB_Name = "modSizePBox"
'      AutoSize picture box 1.01
'  Written by Mike D Sutton of EDais
'
' E-Mail: EDais@mvps.org
' WWW: Http://www.mvps.org/EDais/
'
' Written: 08/11/2002
' Last edited: 18/08/2003
'
'Version history:
' Version 1.01 (17/08/2003):
'   Minor non-impact code changes
'
' Version 1.0 (03/12/2001):
'   ScalePBox() - Stretches the picture stored within a picture
'                 box to fit the bounds of the control
'
'About:
' Simple library for emulating "Stretch" property on picture box
' controls in VB
'
'You use this code at your own risk, I don't accept any
' responsibility for anything nasty it may do to your machine!
'Feel free to re-use this code in your own applications (Yeah,
' like I could stop you anyway ;) However, please don't attempt
' to sell or re-distribute it without my written consent.
'Visit my site for any updates to this an more strange graphics
' related VB code, comments and suggestions always welcome!

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

Private Const COLORONCOLOR As Long = &H3
Private Const HALFTONE As Long = &H4

' This function should be called in he Resize() event of the picture box:
'
' Private Sub Picture1_Resize()
'     Call ScalePBox(Picture1)
' End Sub
'
' Note: The "inInterpolate" parameter will give a better quality scale but with slower performance
'   (This option is apparently not supported by the GDI on Windows 95/98/Me, Untested though..)

Public Function ScalePBox(ByRef inPBox As PictureBox, Optional ByVal inInterpolate As Boolean = True) As Boolean
    Dim DeskWnd As Long, DeskDC As Long
    Dim MyDC As Long, OldDIB As Long
    Dim OldAutoDraw As Boolean
    Dim oldmode As Long
    Dim SrcX As Long, SrcY As Long
    Dim DstX As Long, DstY As Long
    
    If (inPBox.Picture = 0) Then Exit Function ' No picture to stretch!
    
    ' Grab desktop window's DC and create a new DC compatible with it
    DeskWnd = GetDesktopWindow()
    DeskDC = GetDC(DeskWnd)
    MyDC = CreateCompatibleDC(DeskDC)
    Call ReleaseDC(DeskWnd, DeskDC)
    
    If (MyDC) Then ' Hijack picture box's bitmap into temp DC
        OldDIB = SelectObject(MyDC, inPBox.Picture.handle)
        
        If (OldDIB) Then
            With inPBox
                ' Get the scales to draw from and to
                SrcX = .ScaleX(.Picture.Width, vbHimetric, vbPixels)
                SrcY = .ScaleY(.Picture.Height, vbHimetric, vbPixels)
                DstX = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)
                DstY = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels)
                
                ' Remember current auto-redraw state and set to true
                OldAutoDraw = .AutoRedraw
                .AutoRedraw = True
                
                ' Check if the image needs to be re-scaled or not
                If ((SrcX = DstX) And (SrcY = DstY)) Then ' Straight copy
                    Call BitBlt(.hDC, 0, 0, SrcX, SrcY, MyDC, 0, 0, vbSrcCopy)
                Else
                    ' Set the stretch blit mode so it doesn't look terrible on NT systems
                    oldmode = SetStretchBltMode(.hDC, IIf(inInterpolate, HALFTONE, COLORONCOLOR))
                    
                    ' Stretch image
                    Call StretchBlt(.hDC, 0, 0, DstX, DstY, MyDC, 0, 0, SrcX, SrcY, vbSrcCopy)
                    
                    ' Re-set stretch blit mode
                    Call SetStretchBltMode(.hDC, oldmode)
                End If
                
                ' Re-set autoredraw and redraw
                .AutoRedraw = OldAutoDraw
                Call .Refresh
            End With
            
            ' De-select Bitmap object
            Call SelectObject(MyDC, OldDIB)
            
            ' All went well
            ScalePBox = True
        End If
        
        ' Clean up
        Call DeleteDC(MyDC)
    End If
End Function
