Attribute VB_Name = "CommonDialog"
Option Explicit

'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type

Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Const GWL_HINSTANCE = (-6)
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOACTIVATE = &H10
Const HCBT_ACTIVATE = 5
Const WH_CBT = 5

'Const CCHDEVICENAME = 32
'Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&
'Const PD_PRINTSETUP = &H40
'Const PD_DISABLEPRINTTOFILE = &H80000
Const CDERR_DIALOGFAILURE = &HFFFF
Const CDERR_FINDRESFAILURE = &H6
Const CDERR_GENERALCODES = &H0
Const CDERR_INITIALIZATION = &H2
Const CDERR_LOADRESFAILURE = &H7
Const CDERR_LOADSTRFAILURE = &H5
Const CDERR_LOCKRESFAILURE = &H8
Const CDERR_MEMALLOCFAILURE = &H9
Const CDERR_MEMLOCKFAILURE = &HA
Const CDERR_NOHINSTANCE = &H4
Const CDERR_NOHOOK = &HB
Const CDERR_REGISTERMSGFAIL = &HC
Const CDERR_NOTEMPLATE = &H3
Const CDERR_STRUCTSIZE = &H1
Dim hHook As Long

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORS) As Long
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTS) As Long
Declare Function PrintDlgA Lib "comdlg32.dll" (pPrintdlg As PRINTDLGS) As Long
Private Declare Function PageSetupDlgA Lib "comdlg32.dll" (pPagesetupdlg As PageSetupDlg) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 256

Public Const LF_FACESIZE = 32

'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not
'a standard Win32 type.
Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY

Public Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
End Type


Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type PageSetupDlg
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type


Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Type OFNOTIFY
     hdr As NMHDR
     lpOFN As OPENFILENAME
     pszFile As String        '  May be NULL
End Type

Type CHOOSECOLORS
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Type LOGFONT
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
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public Type CHOOSEFONTS
       lStructSize As Long
       hwndOwner As Long          '  caller's window handle
       hdc As Long                '  printer DC/IC or NULL
       lpLogFont As Long          '  ptr. to a LOGFONT struct
       iPointSize As Long         '  10 * size in points of selected font
       flags As Long              '  enum. type flags
       rgbColors As Long          '  returned text color
       lCustData As Long          '  data passed to hook fn.
       lpfnHook As Long           '  ptr. to hook function
       lpTemplateName As String     '  custom template name
       hInstance As Long          '  instance handle of.EXE that
       lpszStyle As String          '  return the style field here
       nFontType As Integer          '  same value reported to the EnumFonts
       MISSING_ALIGNMENT As Integer
       nSizeMin As Long           '  minimum pt size allowed &
       nSizeMax As Long           '  max pt size allowed if
End Type

Type PRINTDLGS
     lStructSize As Long
     hwndOwner As Long
     hDevMode As Long
     hDevNames As Long
     hdc As Long
     flags As Long
     nFromPage As Integer
     nToPage As Integer
     nMinPage As Integer
     nMaxPage As Integer
     nCopies As Integer
     hInstance As Long
     lCustData As Long
     lpfnPrintHook As Long
     lpfnSetupHook As Long
     lpPrintTemplateName As String
     lpSetupTemplateName As String
     hPrintTemplate As Long
     hSetupTemplate As Long
End Type

Public Enum EOrientation
       eoPortrait = 1
       eoLandscape = 2
End Enum

' EPrintQuality constants same as vbPRPQ constants
Public Enum EPrintQuality
      epqDraft = -1
      epqLow = -2
      epqMedium = -3
      epqHigh = -4
End Enum

Public Enum EPageSetupUnits
       epsuInches
       epsuMillimeters
End Enum

Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100

Public Const COLOR_FLAGS = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT

Public Const CF_SCREENFONTS = &H1
Public Const CF_PRINTERFONTS = &H2
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_SHOWHELP = &H4&
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_USESTYLE = &H80&
Public Const CF_EFFECTS = &H100&
Public Const CF_APPLY = &H200&
Public Const CF_ANSIONLY = &H400&
Public Const CF_SCRIPTSONLY = CF_ANSIONLY
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_WYSIWYG = &H8000 '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_SCALABLEONLY = &H20000
Public Const CF_TTONLY = &H40000
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOSIZESEL = &H200000
Public Const CF_SELECTSCRIPT = &H400000
Public Const CF_NOSCRIPTSEL = &H800000
Public Const CF_NOVERTFONTS = &H1000000

Public Const SIMULATED_FONTTYPE = &H8000
Public Const PRINTER_FONTTYPE = &H4000
Public Const SCREEN_FONTTYPE = &H2000
Public Const BOLD_FONTTYPE = &H100
Public Const ITALIC_FONTTYPE = &H200
Public Const REGULAR_FONTTYPE = &H400

Public Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Public Const SHAREVISTRING = "commdlg_ShareViolation"
Public Const FILEOKSTRING = "commdlg_FileNameOK"
Public Const COLOROKSTRING = "commdlg_ColorOK"
Public Const SETRGBSTRING = "commdlg_SetRGBColor"
Public Const HELPMSGSTRING = "commdlg_help"
Public Const FINDMSGSTRING = "commdlg_FindReplace"

Public Const CD_LBSELNOITEMS = -1
Public Const CD_LBSELCHANGE = 0
Public Const CD_LBSELSUB = 1
Public Const CD_LBSELADD = 2


Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Type DevMode
     dmDeviceName As String * CCHDEVICENAME
     dmSpecVersion As Integer
     dmDriverVersion As Integer
     dmSize As Integer
     dmDriverExtra As Integer
     dmFields As Long
     dmOrientation As Integer
     dmPaperSize As Integer
     dmPaperLength As Integer
     dmPaperWidth As Integer
     dmScale As Integer
     dmCopies As Integer
     dmDefaultSource As Integer
     dmPrintQuality As Integer
     dmColor As Integer
     dmDuplex As Integer
     dmYResolution As Integer
     dmTTOption As Integer
     dmCollate As Integer
     dmFormName As String * CCHFORMNAME
     dmUnusedPadding As Integer
     dmBitsPerPel As Integer
     dmPelsWidth As Long
     dmPelsHeight As Long
     dmDisplayFlags As Long
     dmDisplayFrequency As Long
End Type

Public Const PD_ALLPAGES = &H0
Public Const PD_SELECTION = &H1
Public Const PD_PAGENUMS = &H2
Public Const PD_NOSELECTION = &H4
Public Const PD_NOPAGENUMS = &H8
Public Const PD_COLLATE = &H10
Public Const PD_PRINTTOFILE = &H20
Public Const PD_PRINTSETUP = &H40
Public Const PD_NOWARNING = &H80
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNIC = &H200
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_SHOWHELP = &H800
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000

Type DEVNAMES
     wDriverOffset As Integer
     wDeviceOffset As Integer
     wOutputOffset As Integer
     wDefault As Integer
     extra As String * 100
End Type

Public Const DN_DEFAULTPRN = &H1

Public Type SelectedFile
    nFilesSelected As Integer
    'sFiles() As String
    sFile As String
    sLastDirectory As String
    bCanceled As Boolean
End Type

Public Type SelectedColor
    oSelectedColor As OLE_COLOR
    bCanceled As Boolean
End Type

Public Type SelectedFont
    sSelectedFont As String
    bCanceled As Boolean
    bBold As Boolean
    bItalic As Boolean
    nSize As Integer
    bUnderline As Boolean
    bStrikeOut As Boolean
    lColor As Long
    sFaceName As String
End Type

Public Type SelectPrinter
     Printer As DevMode
     bCanceled As Boolean
End Type

Public Type PrinterSetup
       SetupPrinter As PageSetupDlg
       Printer As DevMode
       bCanceled As Boolean
End Type

Public FileDialog As OPENFILENAME
Public ColorDialog As CHOOSECOLORS
Public FontDialog As CHOOSEFONTS
Public PrintDialog As PRINTDLGS
Public Printsetup As PrinterSetup

Dim ParenthWnd As Long

Public Function ShowOpen(ByVal hwnd As Long, Optional ByVal CenterForm As Boolean = True) As SelectedFile
Dim ret As Long
Dim Count As Integer
Dim fileNameHolder As String
Dim LastCharacter As Integer
Dim NewCharacter As Integer
Dim tempFiles(1 To 200) As String
Dim hInst As Long
Dim Thread As Long
    
    ParenthWnd = hwnd
    FileDialog.nStructSize = Len(FileDialog)
    FileDialog.hwndOwner = hwnd
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    FileDialog.sFile = FileDialog.sFile & Space$(2047) & Chr$(0)
    FileDialog.nFileSize = Len(FileDialog.sFile)
    
    'If FileDialog.flags = 0 Then
        FileDialog.flags = OFS_FILE_OPEN_FLAGS
    'End If
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If CenterForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = GetOpenFileName(FileDialog)

    If ret Then
        If Trim$(FileDialog.sFileTitle) = "" Then
            LastCharacter = 0
            Count = 0
            While ShowOpen.nFilesSelected = 0
                NewCharacter = InStr(LastCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare)
                If Count > 0 Then
                    tempFiles(Count) = mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                Else
                    ShowOpen.sLastDirectory = mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                End If
                Count = Count + 1
                If InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare) = InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) Then
                    tempFiles(Count) = mid(FileDialog.sFile, NewCharacter + 1, InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) - NewCharacter - 1)
                    ShowOpen.nFilesSelected = Count
                End If
                LastCharacter = NewCharacter
            Wend
             ShowOpen.sFile = tempFiles(Count)
            'ReDim ShowOpen.sFiles(1 To ShowOpen.nFilesSelected)
            'For Count = 1 To ShowOpen.nFilesSelected
            '    ShowOpen.sFiles(Count) = tempFiles(Count)
            'Next
        Else
            'ReDim ShowOpen.sFiles(1 To 1)
            ShowOpen.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
            ShowOpen.nFilesSelected = 1
            'ShowOpen.sFiles(1) = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
            ShowOpen.sFile = mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
        End If
        ShowOpen.bCanceled = False
        Exit Function
    Else
        ShowOpen.sLastDirectory = ""
        ShowOpen.nFilesSelected = 0
        ShowOpen.bCanceled = True
        ShowOpen.sFile = ""
        'Erase ShowOpen.sFiles
        Exit Function
    End If
End Function

Public Function ShowSave(ByVal hwnd As Long, Optional ByVal CenterForm As Boolean = True) As SelectedFile
Dim ret As Long
Dim hInst As Long
Dim Thread As Long
    
    ParenthWnd = hwnd
    FileDialog.nStructSize = Len(FileDialog)
    FileDialog.hwndOwner = hwnd
    If FileDialog.sFileTitle = "" Then
       FileDialog.sFileTitle = Space$(2048)
    End If
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    If FileDialog.sFile = "" Then
       FileDialog.sFile = Space$(2047) & Chr$(0)
    End If
    FileDialog.nFileSize = Len(FileDialog.sFile)
    
    If FileDialog.flags = 0 Then
        FileDialog.flags = OFS_FILE_SAVE_FLAGS
    End If
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If CenterForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = GetSaveFileName(FileDialog)
    'ReDim ShowSave.sFiles(1)

    If ret Then
        ShowSave.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
        ShowSave.nFilesSelected = 1
        'ShowSave.sFiles(1) = Mid(FileDialog.sFile, 1, FileDialog.nFileOffset - 1)
        'ShowSave.sFiles(1) = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, Chr$(1), vbTextCompare) - FileDialog.nFileOffset - 1)
        ShowSave.sFile = Trim(Replace(FileDialog.sFile, Chr(0), ""))
       ' ShowSave.sFile = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, Chr$(1), vbTextCompare) - FileDialog.nFileOffset - 1)
        ShowSave.bCanceled = False
        Exit Function
    Else
        ErrorCommon
        ShowSave.sLastDirectory = ""
        ShowSave.nFilesSelected = 0
        ShowSave.bCanceled = True
        ShowSave.sFile = ""
        'Erase ShowSave.sFiles
        Exit Function
    End If
End Function

Public Function ShowColor(ByVal hwnd As Long, Optional ByVal CenterForm As Boolean = True) As SelectedColor
Dim customcolors() As Byte  ' dynamic (resizable) array
Dim i As Integer
Dim ret As Long
Dim hInst As Long
Dim Thread As Long

    ParenthWnd = hwnd
    If ColorDialog.lpCustColors = "" Then
        ReDim customcolors(0 To 16 * 4 - 1) As Byte  'resize the array
        For i = LBound(customcolors) To UBound(customcolors)
          customcolors(i) = 254 ' sets all custom colors to white
        Next i
        
        ColorDialog.lpCustColors = StrConv(customcolors, vbUnicode)  ' convert array
    End If
    
    ColorDialog.hwndOwner = hwnd
    ColorDialog.lStructSize = Len(ColorDialog)
    ColorDialog.flags = COLOR_FLAGS
    'ColorDialog.rgbResult = ShowColor.oSelectedColor
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If CenterForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = ChooseColor(ColorDialog)
    If ret Then
        ShowColor.bCanceled = False
        ShowColor.oSelectedColor = ColorDialog.rgbResult
        Exit Function
    Else
        ShowColor.bCanceled = True
        ShowColor.oSelectedColor = &H0&
        Exit Function
    End If
    
End Function

Public Function ShowFont(ByVal hwnd As Long, ByVal startingFontName As String, Optional ByVal CenterForm As Boolean = True) As SelectedFont
Dim ret As Long
Dim lfLogFont As LOGFONT
Dim hInst As Long
Dim Thread As Long
Dim i As Integer
    
    ParenthWnd = hwnd
    FontDialog.nSizeMax = 0
    FontDialog.nSizeMin = 0
    FontDialog.nFontType = Screen.FontCount
    FontDialog.hwndOwner = hwnd
    FontDialog.hdc = 0
    FontDialog.lpfnHook = 0
    FontDialog.lCustData = 0
    FontDialog.lpLogFont = VarPtr(lfLogFont)
    If FontDialog.iPointSize = 0 Then
        FontDialog.iPointSize = 10 * 10
    End If
    FontDialog.lpTemplateName = Space$(2048)
    FontDialog.rgbColors = RGB(0, 255, 255)
    FontDialog.lStructSize = Len(FontDialog)
    
    If FontDialog.flags = 0 Then
        FontDialog.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT 'Or CF_EFFECTS
    End If
    
    For i = 0 To Len(startingFontName) - 1
        lfLogFont.lfFaceName(i) = Asc(mid(startingFontName, i + 1, 1))
    Next
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If CenterForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = ChooseFont(FontDialog)
        
    If ret Then
        ShowFont.bCanceled = False
        ShowFont.bBold = IIf(lfLogFont.lfWeight > 400, 1, 0)
        ShowFont.bItalic = lfLogFont.lfItalic
        ShowFont.bStrikeOut = lfLogFont.lfStrikeOut
        ShowFont.bUnderline = lfLogFont.lfUnderline
        ShowFont.lColor = FontDialog.rgbColors
        ShowFont.nSize = FontDialog.iPointSize / 10
        For i = 0 To 31
            ShowFont.sSelectedFont = ShowFont.sSelectedFont + Chr(lfLogFont.lfFaceName(i))
        Next
    
        ShowFont.sSelectedFont = mid(ShowFont.sSelectedFont, 1, InStr(1, ShowFont.sSelectedFont, Chr(0)) - 1)
        Exit Function
    Else
        ShowFont.bCanceled = True
        Exit Function
    End If
End Function

'Public Function ShowPrinter(ByVal hWnd As Long, Optional ByVal CenterForm As Boolean = True) As Long
'Dim hInst As Long
'Dim Thread As Long
'
'    ParenthWnd = hWnd
'    PrintDialog.hwndOwner = hWnd
'    PrintDialog.lStructSize = Len(PrintDialog)
'
'    'Set up the CBT hook
'    hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
'    Thread = GetCurrentThreadId()
'    If CenterForm = True Then
'        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
'    Else
'        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
'    End If
'
'    ShowPrinter = PrintDlg(PrintDialog)
'
'End Function


Public Function ShowPrinter(Owner As Long, Optional ByVal CenterForm As Boolean = True, Optional PrintFlags As Long) As SelectPrinter

    '-> ShowPrinter Code by Donald Grover

    'Dim PrintDlg As PRINTDLG_TYPE
    Dim DevMode As DevMode ' DEVMODE_TYPE
    Dim DevName As DEVNAMES ' DEVNAMES_TYPE
    Dim lpDevMode As Long
    Dim lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer
    Dim strNewPrinterName As String
    Dim hInst As Long, Thread As Long
    ' Use PrintDialog to get the handle to a memory block with a DevMode and DevName structures

    'Set up the CBT hook
    hInst = GetWindowLong(Owner, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If CenterForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If

    PrintDialog.lStructSize = Len(PrintDialog)
    PrintDialog.hwndOwner = Owner 'Me.hWnd

    PrintDialog.flags = PrintFlags Or PD_DISABLEPRINTTOFILE
    On Error Resume Next

    'Set the current orientation and duplex setting
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DevMode.dmPaperWidth = Printer.Width
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmPaperSize = Printer.PaperSize
    DevMode.dmDuplex = Printer.Duplex
    On Error GoTo 0

    'Allocate memory for the initialization hDevMode structure and copy the settings gathered above into this memory
    PrintDialog.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDialog.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDialog.hDevMode)
    End If

    'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With

    'Allocate memory for the initial hDevName structure and copy the settings gathered above into this memory
    PrintDialog.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDialog.hDevNames)

    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    'Call the print dialog up and let the user make changes
    If PrintDlgA(PrintDialog) <> 0 Then                           'First get the DevName structure.
        lpDevName = GlobalLock(PrintDialog.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDialog.hDevNames

        'Next get the DevMode structure and set the printer properties appropriately
        lpDevMode = GlobalLock(PrintDialog.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDialog.hDevMode)
        GlobalFree PrintDialog.hDevMode
        strNewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> strNewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = strNewPrinterName Then
                    Set Printer = objPrinter
                    'set printer toolbar name at this point
                End If
            Next
        End If

        On Error Resume Next
        'Set printer object properties according to selections made by user
        ShowPrinter.bCanceled = False
        ShowPrinter.Printer = DevMode

        Printer.Copies = DevMode.dmCopies
        Printer.Duplex = DevMode.dmDuplex
        Printer.Orientation = DevMode.dmOrientation
        Printer.PaperSize = DevMode.dmPaperSize
        Printer.PrintQuality = DevMode.dmPrintQuality
        Printer.ColorMode = DevMode.dmColor
        Printer.PaperBin = DevMode.dmDefaultSource

        GlobalFree PrintDialog.hDevMode
        GlobalFree PrintDialog.hDevNames
        On Error GoTo 0
    Else
        ShowPrinter.bCanceled = True
    End If

End Function
'
'Public Function ShowPrinter(Owner As Long, Optional PrintFlags As Long) As pr
'
'    '-> ShowPrinter Code by Donald Grover
'
'    Dim PrintDlg As PRINTDLGS 'PRINTDLG_TYPE
'    Dim DevMode As DevMode '_TYPE
'    Dim DevName As DEVNAMES '_TYPE
'    Dim lpDevMode As Long
'    Dim lpDevName As Long
'    Dim bReturn As Integer
'    Dim objPrinter As Printer
'    Dim NewPrinterName As String
'
'    ' Use PrintDialog to get the handle to a memory block with a DevMode and DevName structures
'
'    PrintDlg.lStructSize = Len(PrintDlg)
'    PrintDlg.hwndOwner = Owner 'Me.hWnd
'
'    PrintDlg.flags = PrintFlags Or PD_DISABLEPRINTTOFILE
'    On Error Resume Next
'
'    'Set the current orientation and duplex setting
'    DevMode.dmDeviceName = Printer.DeviceName
'    DevMode.dmSize = Len(DevMode)
'    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
'    DevMode.dmPaperWidth = Printer.Width
'    DevMode.dmOrientation = Printer.Orientation
'    DevMode.dmPaperSize = Printer.PaperSize
'    DevMode.dmDuplex = Printer.Duplex
'    On Error GoTo 0
'
'    'Allocate memory for the initialization hDevMode structure and copy the settings gathered above into this memory
'    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
'    lpDevMode = GlobalLock(PrintDlg.hDevMode)
'    If lpDevMode > 0 Then
'        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
'        bReturn = GlobalUnlock(PrintDlg.hDevMode)
'    End If
'
'    'Set the current driver, device, and port name strings
'    With DevName
'        .wDriverOffset = 8
'        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
'        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
'        .wDefault = 0
'    End With
'
'    With Printer
'        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
'    End With
'
'    'Allocate memory for the initial hDevName structure and copy the settings gathered above into this memory
'    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
'    lpDevName = GlobalLock(PrintDlg.hDevNames)
'
'    If lpDevName > 0 Then
'        CopyMemory ByVal lpDevName, DevName, Len(DevName)
'        bReturn = GlobalUnlock(lpDevName)
'    End If
'
'    'Call the print dialog up and let the user make changes
'    If PrintDlgA(PrintDlg) <> 0 Then
'        lpDevName = GlobalLock(PrintDlg.hDevNames)
'        CopyMemory DevName, ByVal lpDevName, 45
'        bReturn = GlobalUnlock(lpDevName)
'        GlobalFree PrintDlg.hDevNames
'
'        'Next get the DevMode structure and set the printer properties appropriately
'        lpDevMode = GlobalLock(PrintDlg.hDevMode)
'        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
'        bReturn = GlobalUnlock(PrintDlg.hDevMode)
'        GlobalFree PrintDlg.hDevMode
'        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
'        If Printer.DeviceName <> NewPrinterName Then
'            For Each objPrinter In Printers
'                If UCase$(objPrinter.DeviceName) = NewPrinterName Or InStr(UCase$(objPrinter.DeviceName), NewPrinterName) > 0 Then
'                    Set Printer = objPrinter
'                    'set printer toolbar name at this point
'                End If
'            Next
'        End If
'
'        On Error Resume Next
'        'Set printer object properties according to selections made by user
'        Printer.Copies = DevMode.dmCopies
'        Printer.Duplex = DevMode.dmDuplex
'        Printer.Orientation = DevMode.dmOrientation
'        Printer.PaperSize = DevMode.dmPaperSize
'        Printer.PrintQuality = DevMode.dmPrintQuality
'        Printer.ColorMode = DevMode.dmColor
'        If Printer.PaperBin <> DevMode.dmDefaultSource Then
'        Printer.PaperBin = DevMode.dmDefaultSource
'        End If
'        On Error GoTo 0
'    End If
'
'    ShowPrinter = NewPrinterName
'  '  PrtName$ = NewPrinterName
'
'End Function

Private Function WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rectForm As RECT, rectMsg As RECT
    Dim X As Long, Y As Long
    If lMsg = HCBT_ACTIVATE Then
        'Show the MsgBox at a fixed location (0,0)
        GetWindowRect wParam, rectMsg
        X = Screen.Width / Screen.TwipsPerPixelX / 2 - (rectMsg.Right - rectMsg.Left) / 2
        Y = Screen.Height / Screen.TwipsPerPixelY / 2 - (rectMsg.Bottom - rectMsg.Top) / 2
        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
    End If
    WinProcCenterScreen = False
End Function

Private Function WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rectForm As RECT, rectMsg As RECT
    Dim X As Long, Y As Long
    'On HCBT_ACTIVATE, show the MsgBox centered over Form1
    If lMsg = HCBT_ACTIVATE Then
        'Get the coordinates of the form and the message box so that
        'you can determine where the center of the form is located
        GetWindowRect ParenthWnd, rectForm
        GetWindowRect wParam, rectMsg
        X = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - ((rectMsg.Right - rectMsg.Left) / 2)
        Y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - ((rectMsg.Bottom - rectMsg.Top) / 2)
        'Position the msgbox
        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
     End If
     WinProcCenterForm = False
End Function

Public Function ShowPageSetupDlg(ByVal hwnd As Long, Optional ByVal CenterForm As Boolean = True) As PrinterSetup
                                 ', _
                                 Optional DisableMargins As Boolean, _
                                 Optional DisableOrientation As Boolean, _
                                 Optional DisablePaper As Boolean, _
                                 Optional DisablePrinter As Boolean, _
                                 Optional LeftMargin As Long, _
                                 Optional MinLeftMargin As Long, _
                                 Optional RightMargin As Long, _
                                 Optional MinRightMargin As Long, _
                                 Optional TopMargin As Long, _
                                 Optional MinTopMargin As Long, _
                                 Optional BottomMargin As Long, _
                                 Optional MinBottomMargin As Long, _
                                 Optional Orientation As EOrientation = eoPortrait, _
                                 Optional PrintQuality As EPrintQuality = epqDraft, _
                                 Optional Units As EPageSetupUnits = epsuInches, _
                                 Optional Printer As Object, _
                                 Optional flags As Long) As PrintSetup
    
    Dim m_PSD As PageSetupDlg
    Dim hInst As Long
    Dim Thread As Long

    
    'Set the structure size
    m_PSD.lStructSize = Len(m_PSD)
    'Set the owner window
    m_PSD.hwndOwner = hwnd
    'Set the application instance
    m_PSD.hInstance = App.hInstance
    'no extra flags
    m_PSD.flags = 0
    ParenthWnd = hwnd
     'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If CenterForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    'Show the pagesetup dialog
    If PageSetupDlgA(m_PSD) Then
        'ShowPageSetupDlg = 0
        ' Get DEVMODE structure from PRINTDLG
        Dim dvmode As DevMode, pDevMode As Long
        pDevMode = GlobalLock(m_PSD.hDevMode)
        CopyMemory dvmode, ByVal pDevMode, Len(dvmode)
        Call GlobalUnlock(m_PSD.hDevMode)
        'PaperSize = dvmode.dmPaperSize
        'Orientation = dvmode.dmOrientation
        'PrintQuality = dvmode.dmPrintQuality
        
        'LeftMargin = m_PSD.lStructSize
        'RightMargin = 0 'm_PSD.rtMinMargin.Right
        'TopMargin = 0
        'BottomMargin = 0
        'MinLeftMargin = m_PSD.rtMinMargin.Left / 100
        'MinRightMargin = m_PSD.rtMinMargin.Right / 100
        'MinTopMargin = m_PSD.rtMinMargin.Top / 100
        'MinBottomMargin = m_PSD.rtMinMargin.Bottom / 100
        ShowPageSetupDlg.SetupPrinter = m_PSD
        ShowPageSetupDlg.Printer = dvmode
        ' Set default printer properties
        On Error Resume Next
        If Not (Printer Is Nothing) Then
            Printer.Copies = dvmode.dmCopies
            Printer.Orientation = dvmode.dmOrientation
            Printer.PaperSize = dvmode.dmPaperSize
            Printer.PrintQuality = dvmode.dmPrintQuality
        End If
    Else
       ' ShowPageSetupDlg = -1
    End If
    
End Function


Function ErrorCommon() 'as strring
        Select Case CommDlgExtendedError
        Case CDERR_DIALOGFAILURE
            MsgBox "The dialog box could not be created."
        Case CDERR_FINDRESFAILURE
            MsgBox "The common dialog box function failed to find a specified resource."
        Case CDERR_INITIALIZATION
            MsgBox "The common dialog box function failed during initialization."
        Case CDERR_LOADRESFAILURE
            MsgBox "The common dialog box function failed to load a specified resource."
        Case CDERR_LOADSTRFAILURE
            MsgBox "The common dialog box function failed to load a specified string."
        Case CDERR_LOCKRESFAILURE
            MsgBox "The common dialog box function failed to lock a specified resource."
        Case CDERR_MEMALLOCFAILURE
            MsgBox "The common dialog box function was unable to allocate memory for internal structures."
        Case CDERR_MEMLOCKFAILURE
            MsgBox "The common dialog box function was unable to lock the memory associated with a handle."
        Case CDERR_NOHINSTANCE
            MsgBox "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding instance handle."
        Case CDERR_NOHOOK
            MsgBox "The ENABLEHOOK flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a pointer to a corresponding hook procedure."
        Case CDERR_REGISTERMSGFAIL
            MsgBox "The RegisterWindowMessage function returned an error code when it was called by the common dialog box function."
        Case CDERR_NOTEMPLATE
            MsgBox "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding template."
        Case CDERR_STRUCTSIZE
            MsgBox "The lStructSize member of the initialization structure for the corresponding common dialog box is invalid."
        Case Else
            MsgBox "Undefined error ..."
    End Select

End Function
