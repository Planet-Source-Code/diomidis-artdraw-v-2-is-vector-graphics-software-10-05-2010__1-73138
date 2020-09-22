VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSectionRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==================================================================
' FileName:    cDIBSectionRegion.cls
' Author:      Steve McMahon
'
' Converts a cDIBSection object into a region which you can apply
' to a form, UserControl or PictureBox (in fact, anything with a
' hWnd property).
'
' Also includes functions to Save a region to a file, and to Load
' a region either from a file or from a resource.  The resource
' loading code is useful because it demonstrates how to load
' arbitrary resource data from any external library.
'
' ------------------------------------------------------------------
' Visit vbAccelerator - advanced, hardcore VB with full source code
' http://vbaccelerator.com/
' mailto:steve@vbaccelerator.com
'
' ==================================================================

' API for creating a region:
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_DIFF = 4
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long


' API for saving and loading a region:
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
' API for getting data from an external library module:
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, lpName As Any, lpType As Any) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Const LOAD_LIBRARY_AS_DATAFILE = &H2&
Private Const RT_RCDATA = 10&

' API for reading cDIBSection bits:
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

' Implementation:
Private m_hRgn As Long
Private m_hWnd() As Long
Private m_iCount As Long

Public Property Get Applied(ByVal hwnd As Long) As Boolean
   Applied = Not (plIndex(hwnd) = 0)
End Property
Public Property Let Applied(ByVal hwnd As Long, ByVal bState As Boolean)
Dim i As Long
Dim lIndex As Long
   lIndex = plIndex(hwnd)
   If bState Then
      If (lIndex = 0) Then
         ' Apply to window:
         m_iCount = m_iCount + 1
         ReDim Preserve m_hWnd(1 To m_iCount) As Long
         m_hWnd(m_iCount) = hwnd
         SetWindowRgn m_hWnd(m_iCount), m_hRgn, True
      Else
         ' already applied, reset apply state jic
         SetWindowRgn m_hWnd(m_iCount), m_hRgn, True
      End If
   Else
      If (lIndex = 0) Then
         ' Not applied, reset state jic
         SetWindowRgn hwnd, 0, True
      Else
         ' Applied, reset:
         SetWindowRgn hwnd, 0, True
         If m_iCount > 1 Then
            For i = lIndex To m_iCount - 1
               m_hWnd(i) = m_hWnd(i + 1)
            Next i
            m_iCount = m_iCount - 1
            ReDim Preserve m_hWnd(1 To m_iCount) As Long
         Else
            m_iCount = 0
            Erase m_hWnd
         End If
      End If
   End If
End Property
Private Property Get plIndex(ByVal hwnd As Long) As Long
Dim i As Long
Dim lIndex As Long
   For i = 1 To m_iCount
      If hwnd = m_hWnd(i) Then
         plIndex = i
         Exit For
      End If
   Next i
End Property
Public Property Get AppliedToCount() As Long
   AppliedToCount = m_iCount
End Property
Public Property Get hWndForIndex(ByVal lIndex As Long) As Long
   hWndForIndex = m_hWnd(lIndex)
End Property

Private Sub UnApply()
Dim i As Long
   For i = 1 To m_iCount
      If Not m_hWnd(i) = 0 Then
         SetWindowRgn m_hWnd(i), 0, True
         m_hWnd(i) = 0
      End If
   Next i
   m_iCount = 0
End Sub
Public Sub Destroy()
   UnApply
   If Not m_hRgn = 0 Then
      DeleteObject m_hRgn
   End If
   m_hRgn = 0
End Sub

Public Sub Create( _
      ByRef cDib As cDIBSection, _
      Optional ByRef lTransColor As Long = 0 _
   )
Dim X As Long, Y As Long
Dim lX As Long
Dim yStart As Long
Dim bStart As Boolean
Dim hRgnTemp As Long
Dim bR As Byte, bG As Byte, bB As Byte
Dim lWidth As Long, lHeight As Long
Dim bDib() As Byte
Dim tSA As SAFEARRAY2D

   Destroy
   
   ' The transparent colour:
   bR = (lTransColor And &HFF&)
   bG = (lTransColor And &HFF00&) \ &H100&
   bB = (lTransColor And &HFF0000) \ &H10000
   
   ' Create the base region
   m_hRgn = CreateRectRgn(0, 0, cDib.Width, cDib.Height)
   Debug.Assert (m_hRgn <> 0)
   If m_hRgn <> 0 Then
      ' Get the DIB into byte array:
      With tSA
          .cbElements = 1
          .cDims = 2
          .Bounds(0).lLbound = 0
          .Bounds(0).cElements = cDib.Height
          .Bounds(1).lLbound = 0
          .Bounds(1).cElements = cDib.BytesPerScanLine()
          .pvData = cDib.DIBSectionBitsPtr
      End With
      CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
               
      lWidth = cDib.BytesPerScanLine \ 3
      lHeight = cDib.Height
      For X = 0 To (lWidth - 1) * 3 Step 3
         ' DIB Sections are "upside down" :)
         For Y = lHeight - 1 To 0 Step -1
            If bDib(X, Y) = bB And bDib(X + 1, Y) = bG And bDib(X + 2, Y) = bR Then
               If Not bStart Then
                  yStart = lHeight - 1 - Y
                  bStart = True
               End If
            Else
               If bStart Then
                  hRgnTemp = CreateRectRgn(lX, yStart, lX + 1, lHeight - 1 - Y)
                  CombineRgn m_hRgn, hRgnTemp, m_hRgn, RGN_XOR
                  DeleteObject hRgnTemp
                  bStart = False
               End If
            End If
         Next Y
         If bStart Then
            hRgnTemp = CreateRectRgn(lX, yStart, lX + 1, lHeight - 1 - Y)
            CombineRgn m_hRgn, hRgnTemp, m_hRgn, RGN_XOR
            DeleteObject hRgnTemp
            bStart = False
         End If
         lX = lX + 1
      Next X
      
      CopyMemory ByVal VarPtrArray(bDib), 0&, 4
      
   End If
End Sub

Public Function Save(ByVal sPath As String) As Boolean
Dim iFile As Long
Dim nBytes As Long
Dim b() As Byte

On Error GoTo ErrorHandler ' Out of memory

   If Not m_hRgn = 0 Then
      
      nBytes = GetRegionData(m_hRgn, 0, ByVal 0&)
      If nBytes > 0 Then
         ReDim b(0 To nBytes - 1) As Byte
         If nBytes = GetRegionData(m_hRgn, nBytes, b(0)) Then
            On Error Resume Next ' Attempt to kill file
            Kill sPath
            On Error GoTo ErrorHandler ' Error handler checks for file error
            iFile = FreeFile
            Open sPath For Binary Access Write Lock Read As #iFile
            Put #iFile, , b
            Close #iFile
            Save = True
         Else
            Err.Raise 26012, App.EXEName & ".cDIBSectionRegion", "Unable to get region data"
         End If
      Else
         Err.Raise 26011, App.EXEName & ".cDIBSectionRegion", "Unable to determine size of region"
      End If
   Else
      Err.Raise 26010, App.EXEName & ".cDIBSectionRegion", "No region to save"
   End If
   Exit Function
   
ErrorHandler:
Dim lErr As Long, sErr As String
   lErr = Err.Number: sErr = Err.Description
   If iFile > 0 Then
      Close #iFile
   End If
   Err.Raise lErr, App.EXEName & ".cDIBSectionRegion", sErr
   Exit Function
End Function

Public Function LoadFromFile(ByVal sFileName As String) As Boolean
Dim iFile As Long
Dim b() As Byte
On Error GoTo ErrorHandler

   iFile = FreeFile
   Open sFileName For Binary Access Read Lock Write As #iFile
   ReDim b(0 To LOF(iFile) - 1) As Byte
   Get #iFile, , b
   Close #iFile

   LoadFromFile = pbLoadFromByteArray(b())
   Exit Function

ErrorHandler:
Dim lErr As Long, sErr As String
   lErr = Err.Number: sErr = Err.Description
   If iFile > 0 Then
      Close #iFile
   End If
   Err.Raise lErr, App.EXEName & ".cDIBSectionRegion", sErr
   Exit Function
End Function

Public Function LoadFromResource(ByVal vID As Variant, Optional ByVal sDLL As String = "") As Boolean
Dim b() As Byte


   If sDLL = "" Then
      ' Local data
      b = LoadResData(vID, 10)
      LoadFromResource = pbLoadFromByteArray(b())
   Else
      Dim hMod As Long, hRes As Long, hGlobal As Long, lPtr As Long, lSize As Long
      Dim lId As Long, sID As String, lR As Long
      
      ' Load from external module, for data only:
      hMod = LoadLibraryEx(sDLL, ByVal 0&, LOAD_LIBRARY_AS_DATAFILE)
      If Not hMod = 0 Then
         If IsNumeric(vID) Then
            sID = "#" & CStr(vID)
         End If
         hRes = FindResource(hMod, ByVal sID, ByVal RT_RCDATA)
         If Not hRes = 0 Then
            lSize = SizeofResource(hMod, hRes)
            hGlobal = LoadResource(hMod, hRes)
            If Not hGlobal = 0 Then
               lPtr = LockResource(hGlobal)
               If Not lPtr = 0 Then
                  ReDim b(0 To lSize - 1) As Byte
                  CopyMemory b(0), ByVal lPtr, lSize
                  LoadFromResource = pbLoadFromByteArray(b())
               End If
            Else
               Err.Raise 26014, App.EXEName & ".cDIBSectionRegion", "Cannot access data for resource with ID " & vID & " could not be found"
            End If
         Else
            Err.Raise 26014, App.EXEName & ".cDIBSectionRegion", "Resource with ID " & vID & " could not be found"
         End If
         lR = FreeLibrary(hMod)
         Debug.Assert Not (lR = 0)
         If Not lR = 0 Then
            hMod = 0
         End If
      Else
         Err.Raise 26013, App.EXEName & ".cDIBSectionRegion", "Can't open DLL for Resource Access"
      End If
   End If
   Exit Function

ErrorHandler:
Dim lErr As Long, sErr As String
   lErr = Err.Number: sErr = Err.Description
   If Not hMod = 0 Then
      lR = FreeLibrary(hMod)
      Debug.Assert Not (lR = 0)
   End If
   Err.Raise lErr, App.EXEName & ".cDIBSectionRegion", sErr
   Exit Function
End Function

Private Function pbLoadFromByteArray(ByRef b() As Byte) As Boolean
Dim dwCount As Long
   
   Destroy
   dwCount = UBound(b) - LBound(b) + 1
   m_hRgn = ExtCreateRegion(ByVal 0&, dwCount, b(0))
   pbLoadFromByteArray = Not (m_hRgn = 0)
   
End Function

Private Sub Class_Terminate()
   Destroy
End Sub


