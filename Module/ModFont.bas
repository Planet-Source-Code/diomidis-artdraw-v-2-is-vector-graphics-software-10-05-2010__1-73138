Attribute VB_Name = "ModFonts"
Option Explicit


'Font enumeration types
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte 'OR STRING *33

        lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte

        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type

' ntmFlags field flags
Private Const NTM_REGULAR = &H40&
Private Const NTM_BOLD = &H20&
Private Const NTM_ITALIC = &H1&

'  tmPitchAndFamily flags
Private Const TMPF_FIXED_PITCH = &H1

Private Const TMPF_VECTOR = &H2
Private Const TMPF_DEVICE = &H8
Private Const TMPF_TRUETYPE = &H4

Private Const ELF_VERSION = 0
Private Const ELF_CULTURE_LATIN = 0

'  EnumFonts Masks
Private Const RASTER_FONTTYPE = &H1
Private Const DEVICE_FONTTYPE = &H2
'Private Const TRUETYPE_FONTTYPE = &H4

Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" _
                            (ByVal hDC As Long, ByVal lpszFamily As String, _
                            ByVal lpEnumFontFamProc As Long, lParam As Any) As Long

'Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFont Lib "gdi32.dll" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetPathAPI Lib "gdi32.dll" Alias "GetPath" (ByVal hDC As Long, ByRef lpPoints As Any, ByRef lpTypes As Any, ByVal nSize As Long) As Long
Private Declare Function GetPath Lib "gdi32" (ByVal hDC As Long, lpPoint As PointAPI, lpTypes As Byte, ByVal nSize As Long) As Long
Private Declare Function PolyBezierTo Lib "gdi32" (ByVal hDC As Long, lpPt As PointAPI, ByVal cCount As Long) As Long
Private Declare Function PolyDraw Lib "gdi32" (ByVal hDC As Long, lpPt As PointAPI, lpbTypes As Byte, ByVal cCount As Long) As Long

Private Declare Function FillPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Dim m_PointCoords() As POINTAPI
Dim m_PointTypes() As Byte
Dim m_NumPoints As Long

Private Type PointAPI
    x As Long
    y As Long
End Type

Private Const LOGPIXELSY = 90                    'For GetDeviceCaps - returns the height of a logical pixel
'Private Const ANSI_CHARSET = 0                   'Use the default Character set
'Private Const CLIP_LH_ANGLES = 16                ' Needed for tilted fonts.
'Private Const OUT_TT_PRECIS = 9                  'Tell it to use True Types when Possible
'Private Const PROOF_QUALITY = 9                  'Make it as clean as possible.
'Private Const DEFAULT_PITCH = 0                  'We want the font to take whatever pitch it defaults to
'Private Const FF_DONTCARE = 0                    'Use whatever fontface it is.

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

'drawtext
Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000
Public Const DT_EDITCONTROL = &H2000
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000
Public Const DT_MODIFYSTRING = &H10000
Public Const DT_RTLREADING = &H20000
Public Const DT_WORD_ELLIPSIS = &H40000
Public Const DC_GRADIENT = &H20

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, ByVal lpRect As Any, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Public Const ETO_OPAQUE = 2
' Font weight constants.
Private Const FW_DONTCARE = 0
Enum FontWeight
    FW_THIN = 100
    FW_EXTRALIGHT = 200
    FW_LIGHT = 300
    FW_NORMAL = 400
    FW_MEDIUM = 500
    FW_SEMIBOLD = 600
    FW_BOLD = 700
    FW_EXTRABOLD = 800
    FW_HEAVY = 900
    FW_ULTRALIGHT = FW_EXTRALIGHT
    FW_REGULAR = FW_NORMAL
    FW_DEMIBOLD = FW_SEMIBOLD
    FW_ULTRABOLD = FW_EXTRABOLD
End Enum
Private Const FW_BLACK = FW_HEAVY

' Character set constants.
Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const OEM_CHARSET = 255

' Output precision constants.
Private Const OUT_CHARACTER_PRECIS = 2
Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_DEVICE_PRECIS = 5
Private Const OUT_RASTER_PRECIS = 6
Private Const OUT_STRING_PRECIS = 1
Private Const OUT_STROKE_PRECIS = 3
Private Const OUT_TT_ONLY_PRECIS = 7
Private Const OUT_TT_PRECIS = 4

' Clipping precision constants.
Private Const CLIP_CHARACTER_PRECIS = 1
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_EMBEDDED = &H80
Private Const CLIP_LH_ANGLES = &H10
Private Const CLIP_STROKE_PRECIS = 2
Private Const CLIP_TO_PATH = 4097
Private Const CLIP_TT_ALWAYS = &H20

' Character quality constants.
Private Const DEFAULT_QUALITY = 0
Private Const DRAFT_QUALITY = 1
Private Const PROOF_QUALITY = 2

' Pitch and family constants.
Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2
Private Const TRUETYPE_FONTTYPE = &H4
Private Const FF_DECORATIVE = 80  '  Old English, etc.
Private Const FF_DONTCARE = 0     '  Don't care or don't know.
Private Const FF_MODERN = 48      '  Constant stroke width, serifed or sans-serifed.
Private Const FF_ROMAN = 16       '  Variable stroke width, serifed.
Private Const FF_SCRIPT = 64      '  Cursive, etc.
Private Const FF_SWISS = 32

'Draw a rotated string centered at the indicated position using the indicated font parameters.
Public Sub CenterText(ByVal pic As PictureBox, _
                      ByVal xmid As Single, ByVal ymid As Single, _
                      ByVal txt As String, _
                      ByVal nHeight As Long, _
                      Optional ByVal nWidth As Long = 0, _
                      Optional ByVal nEscapement As Long = 0, _
                      Optional ByVal fnWeight As Long = FW_NORMAL, _
                      Optional ByVal fbItalic As Long = False, _
                      Optional ByVal fbUnderline As Long = False, _
                      Optional ByVal fbStrikeOut As Long = False, _
                      Optional ByVal fbCharSet As Long = DEFAULT_CHARSET, _
                      Optional ByVal fbOutputPrecision As Long = OUT_TT_ONLY_PRECIS, _
                      Optional ByVal fbClipPrecision As Long = CLIP_EMBEDDED, _
                      Optional ByVal fbQuality As Long = DEFAULT_QUALITY, _
                      Optional ByVal fbPitchAndFamily As Long = TRUETYPE_FONTTYPE, _
                      Optional ByVal lpszFace As String = "Arial", _
                      Optional ByRef FWidth As Single, Optional ByRef FHeight As Single)

Dim NewFont As Long
Dim oldfont As Long
Dim text_metrics As TEXTMETRIC
Dim internal_leading As Single
Dim total_hgt As Single
Dim text_wid As Long
Dim text_hgt As Single
Dim text_bound_wid As Single
Dim text_bound_hgt As Single
Dim total_bound_wid As Single
Dim total_bound_hgt As Single
Dim theta As Single
Dim phi As Single
Dim X1 As Single
Dim Y1 As Single
Dim X2 As Single
Dim Y2 As Single
Dim X3 As Single
Dim Y3 As Single
Dim X4 As Single
Dim Y4 As Single
Dim NRECT As RECT
Dim flags As Long
    
    ' Create the font.
    NewFont = CreateFont(nHeight, nWidth, nEscapement, 0, fnWeight, fbItalic, _
                         fbUnderline, fbStrikeOut, fbCharSet, fbOutputPrecision, _
                         fbClipPrecision, fbQuality, fbPitchAndFamily, lpszFace)
                         
    oldfont = SelectObject(pic.hDC, NewFont)

    ' Get the font metrics.
    GetTextMetrics pic.hDC, text_metrics
    internal_leading = pic.ScaleY(text_metrics.tmInternalLeading, vbPixels, pic.ScaleMode)
    total_hgt = pic.ScaleY(text_metrics.tmHeight, vbPixels, pic.ScaleMode)
    text_hgt = total_hgt - internal_leading
    text_wid = 0
    text_wid = CLng(pic.TextWidth(txt))
    
    FWidth = text_wid
    FHeight = text_hgt
    
    ' Get the bounding box geometry.
    theta = nEscapement / 10 / 180 * Pi
    phi = Pi / 2 - theta
    text_bound_wid = text_hgt * Cos(phi) + text_wid * Cos(theta)
    text_bound_hgt = text_hgt * Sin(phi) + text_wid * Sin(theta)
    total_bound_wid = total_hgt * Cos(phi) + text_wid * Cos(theta)
    total_bound_hgt = total_hgt * Sin(phi) + text_wid * Sin(theta)

    ' Find the desired center point.
    X1 = xmid
    Y1 = ymid

    ' Subtract half the height and width of the text
    ' bounding box. This puts (x1, y2) in the upper
    ' left corner of the text bounding box.
    X1 = X1 - text_bound_wid / 2
    Y1 = Y1 - text_bound_hgt / 2

    ' The start position's X coordinate belongs at
    ' the left edge of the text bounding box, so
    ' x1 is correct. Move the Y coordinate down to
    ' its start position.
    Y1 = Y1 + text_wid * Sin(theta)

    ' Move (x1, y1) to the start corner of the outer bounding box.
    X1 = X1 - (total_bound_wid - text_bound_wid)
    Y1 = Y1 - (total_bound_hgt - text_bound_hgt)
   
    TextOut pic.hDC, X1, Y1, txt, Len(txt)
    
    ' Reselect the old font and delete the new one.
    NewFont = SelectObject(pic.hDC, oldfont)
    Call DeleteObject(NewFont)
        
End Sub

Public Function CreateFontWmf(ByVal hDC As Long, _
                       ByVal nHeight As Long, _
                       Optional ByVal nWidth As Long = 0, _
                       Optional ByVal nEscapement As Long = 0, _
                       Optional ByVal fnWeight As Long = FW_NORMAL, _
                       Optional ByVal fbItalic As Long = False, _
                       Optional ByVal fbUnderline As Long = False, _
                       Optional ByVal fbStrikeOut As Long = False, _
                       Optional ByVal fbCharSet As Long = DEFAULT_CHARSET, _
                       Optional ByVal fbOutputPrecision As Long = OUT_TT_ONLY_PRECIS, _
                       Optional ByVal fbClipPrecision As Long = CLIP_EMBEDDED, _
                       Optional ByVal fbQuality As Long = PROOF_QUALITY, _
                       Optional ByVal fbPitchAndFamily As Long = TRUETYPE_FONTTYPE, _
                       Optional ByVal lpszFace As String = "Arial") As Long

    ' Create the font.
    CreateFontWmf = CreateFont(nHeight, nWidth, nEscapement, 0, fnWeight, fbItalic, fbUnderline, fbStrikeOut, fbCharSet, fbOutputPrecision, fbClipPrecision, fbQuality, fbPitchAndFamily, lpszFace)
            
End Function


'Read Path text and make PointCoolds and Type for draw
Public Sub ReadPathText(ByVal Obj As PictureBox, _
                        ByVal txt As String, _
                        ByRef Point_Coords() As PointAPI, _
                        ByRef Point_Types() As Byte, _
                        ByVal NumPoints As Long)
    Dim ret As Long
    ret = BeginPath(Obj.hDC)
    Obj.Print txt
    ret = EndPath(Obj.hDC)
    NumPoints = 0
    NumPoints = GetPathAPI(Obj.hDC, ByVal 0&, ByVal 0&, 0)

    If (NumPoints) Then
        ReDim Point_Coords(NumPoints - 1)
        ReDim Point_Types(NumPoints - 1)
        'Get the path data from the DC
        Call GetPathAPI(Obj.hDC, Point_Coords(0), Point_Types(0), NumPoints)
    End If

End Sub

Public Sub LoadFonts(ByVal ComboBox As Object)
    Dim hDC As Long
    ComboBox.Clear
    hDC = GetDC(ComboBox.hWnd)
    Call EnumFontFamilies(hDC, vbNullString, AddressOf EnumFontFamProc, ComboBox)
    Call ReleaseDC(ComboBox.hWnd, hDC)
End Sub

Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
    ByVal FontType As Long, lParam As ComboBox) As Long
    On Local Error Resume Next
    Dim FaceName As String
    Dim FullName As String
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    Call lParam.AddItem(Left$(FaceName, InStr(FaceName, vbNullChar) - 1))
    EnumFontFamProc = 1
End Function

Public Sub Draw_Example(pic As PictureBox, txt As String, _
                         m3Deffect As Boolean, mRaised As Boolean, _
                         mColor1 As Long, mColor2 As Long, _
                         Optional m3DAngle As Single = 0, Optional mAngle As Single = 0)
    Dim XX As Single, yy As Single, Ang As Single, a As Integer, x As Single, y As Single

    XX = pic.CurrentX
    yy = pic.CurrentY
    If mRaised = True Then GoTo raisedd
    If m3Deffect = False Then GoTo ddd1
    
hh1:
    If mAngle > 360 Then mAngle = mAngle - 360: GoTo hh1
    If mAngle <= 0 Then mAngle = 0: GoTo hhjh
    Ang = 360 / mAngle
    Ang = ((Pi * 2) / Ang) + (Pi / 2)
hhjh:
    For a = 1 To m3DAngle
        pic.ForeColor = mColor1
        pic.DrawWidth = 1
        pic.Line (XX + a * Cos(Ang), yy + a * Sin(Ang))-(XX + a * Cos(Ang), yy + a * Sin(Ang)), mColor1
        pic.Print txt
    Next a
    GoTo ddff1
raisedd:
    For x = XX - 1 To XX Step 1
        For y = yy - 1 To yy Step 1
            pic.ForeColor = RGB(255, 255, 255)
            pic.DrawWidth = 1: pic.Line (x, y)-(x, y), RGB(255, 255, 255): pic.Print txt
    Next y, x
   For x = XX To XX + 1 Step 1
       For y = yy To yy + 1 Step 1
           pic.ForeColor = RGB(0, 0, 0)
           pic.DrawWidth = 1: pic.Line (x, y)-(x, y), RGB(0, 0, 0): pic.Print txt
    Next y, x
GoTo ddff1
ddd1:

   For x = XX - 1 To XX + 1
       For y = yy - 1 To yy + 1
           pic.ForeColor = mColor1
           pic.DrawWidth = 1: pic.Line (x, y)-(x, y), mColor1: pic.Print txt
   Next y, x
ddff1:
    pic.ForeColor = mColor2
    pic.DrawWidth = 1: pic.Line (XX, yy)-(XX, yy), mColor2: pic.Print txt
    
 End Sub

'dIST 1-100, Pxx=-1 - 1 ,Pyy=-1 - 1
Public Sub Print3D(Ob As Object, txt As String, Dist As Integer, _
                   mColor1 As Long, mColor2 As Long, _
                   Pxx As Single, PYY As Single)
                   
On Error Resume Next
Dim Sr As Single, sg As Single, Sb As Single
Dim r1 As Long, g1 As Long, b1 As Long
Dim r2 As Long, g2 As Long, b2 As Long
Dim TMidX As Long, TMidY As Long, XX As Long, yy As Long
   ' Ob.Cls
    'do 3D
    Pxx = Pxx / 10
    PYY = PYY / 10
    
    'make sure text is always centerred
    TMidX = (Ob.Width / 2) - (Ob.TextWidth(txt$) / 2)
    TMidY = (Ob.Height / 2) - (Ob.TextHeight(txt$) / 2)
    TMidX = TMidX - ((Pxx * Dist) / 2)
    TMidY = TMidY - ((PYY * Dist) / 2)
    SplitRGB mColor1, r1, g1, b1
    SplitRGB mColor2, r2, g2, b2
    
    Sr = (r2 - r1) / Dist
    sg = (g2 - g1) / Dist
    Sb = (b2 - b1) / Dist
    
    'print a lot of text
    For XX = 0 To Dist - 1
        Ob.CurrentX = TMidX + (XX * Pxx)
        Ob.CurrentY = TMidY + (XX * PYY)
        r1 = r1 + Sr
        g1 = g1 + sg
        b1 = b1 + Sb
        'the values cannot be < 0
        If Int(r1) < 0 Then r1 = 0
        If Int(g1) < 0 Then g1 = 0
        If Int(b1) < 0 Then b1 = 0
        Ob.ForeColor = RGB(r1, g1, b1)
        Ob.Print txt
    Next XX
    

End Sub

