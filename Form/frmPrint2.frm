VERSION 5.00
Begin VB.Form frmPrint 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Print"
   ClientHeight    =   5820
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7260
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   388
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optPaperSize 
      Caption         =   "A4"
      Height          =   195
      Index           =   0
      Left            =   5715
      TabIndex        =   17
      Top             =   2955
      Width           =   990
   End
   Begin VB.OptionButton optPaperSize 
      Caption         =   "8.5 x 11"""
      Height          =   195
      Index           =   1
      Left            =   5700
      TabIndex        =   16
      Top             =   3270
      Width           =   1080
   End
   Begin VB.Frame FrmPaper 
      Caption         =   "Paper size"
      Height          =   975
      Left            =   5280
      TabIndex        =   15
      Top             =   2640
      Width           =   1875
   End
   Begin VB.Frame FrmSize 
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   5325
      TabIndex        =   10
      Top             =   45
      Width           =   1785
      Begin VB.HScrollBar scrPercent 
         Height          =   240
         LargeChange     =   10
         Left            =   165
         Max             =   400
         Min             =   10
         TabIndex        =   11
         Top             =   900
         Value           =   10
         Width           =   1410
      End
      Begin VB.Label LabWH 
         Alignment       =   2  'Center
         Caption         =   "W x H"
         ForeColor       =   &H80000007&
         Height          =   225
         Left            =   240
         TabIndex        =   14
         Top             =   330
         Width           =   1290
      End
      Begin VB.Label Label2 
         Caption         =   "Scale"
         Height          =   240
         Left            =   135
         TabIndex        =   13
         Top             =   630
         Width           =   450
      End
      Begin VB.Label LabPercent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   600
         TabIndex        =   12
         Top             =   615
         Width           =   720
      End
   End
   Begin VB.Frame frmOrientation 
      Caption         =   "Orientation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5295
      TabIndex        =   6
      Top             =   1485
      Width           =   1860
      Begin VB.PictureBox picOrientation 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   30
         ScaleHeight     =   675
         ScaleWidth      =   1785
         TabIndex        =   7
         Top             =   300
         Width           =   1785
         Begin VB.OptionButton optOrient 
            Caption         =   "Portrait"
            Height          =   255
            Index           =   0
            Left            =   570
            TabIndex        =   9
            Top             =   15
            Width           =   915
         End
         Begin VB.OptionButton optOrient 
            Caption         =   "Landscape"
            Height          =   255
            Index           =   1
            Left            =   570
            TabIndex        =   8
            Top             =   345
            Width           =   1110
         End
         Begin VB.Image imgPrinterOrien 
            Height          =   465
            Left            =   90
            Top             =   90
            Width           =   345
         End
         Begin VB.Image imgPage 
            Height          =   465
            Index           =   0
            Left            =   75
            Picture         =   "frmPrint2.frx":0000
            Top             =   675
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Image imgPage 
            Height          =   345
            Index           =   1
            Left            =   330
            Picture         =   "frmPrint2.frx":05AE
            Top             =   705
            Visible         =   0   'False
            Width           =   465
         End
      End
   End
   Begin VB.PictureBox picIN 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   8130
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   5
      Top             =   180
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   480
      Left            =   5595
      TabIndex        =   4
      Top             =   4575
      Width           =   1425
   End
   Begin VB.PictureBox picPaper 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   1290
      ScaleHeight     =   295
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   208
      TabIndex        =   1
      Top             =   555
      Width           =   3150
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1950
         Left            =   420
         MousePointer    =   15  'Size All
         ScaleHeight     =   130
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   146
         TabIndex        =   2
         Top             =   735
         Width           =   2190
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Default         =   -1  'True
      Height          =   480
      Left            =   5610
      TabIndex        =   0
      Top             =   3960
      Width           =   1425
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   5265
      Left            =   105
      Top             =   135
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   5265
      Left            =   90
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "The image can be moved with the mouse"
      Height          =   240
      Left            =   1215
      TabIndex        =   3
      Top             =   5475
      Width           =   3345
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmPrint.frm

Option Explicit

Private Multiplier As Single
Private aMouseDown As Boolean
Private xp1 As Long, yp1 As Long
Private aOpt As Boolean
Private PaperType As Integer  ' 0 A4, 1 USA
Private Percent As Long
Private StepW As Single, StepH As Single  ' pixels per inch or cm step
Private paperW As Long, paperH As Long
Private Canceled As Boolean
Private FullHeight As Long, FullWidth As Long
Private iDATA() As Byte           'holds bitmap data
Private DIBInfo As BITMAPINFO      'Device Ind. Bitmap info structure

Private Sub cmdPrint_Click()
   Me.Hide
End Sub

Private Sub PrintPicture(Prn As Printer, picPaper As PictureBox, pic As PictureBox, picIN As PictureBox)
   Const vbHiMetric As Integer = 8
   Const vbTwips As Integer = 1
   Const vbPixels As Integer = 3
   Dim FullPrnWidth As Double
   Dim FullPrnHeight As Double
   Dim PrnPicWidth As Double
   Dim PrnPicHeight As Double
   Dim PicHeight As Double
   Dim PrnPicLeft As Double
   Dim PrnPicTop As Double

   Dim PrnWidth As Double
   Dim PrnHeight As Double
   Dim PrnLeft As Double
   Dim PrnTop As Double
   
 
   ' Calculate the dimensions of the printable area in HiMetric.
   FullPrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
   FullPrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
   
   ' Adjust printable width & height
   PrnWidth = FullPrnWidth * pic.Width / picPaper.Width
   PrnHeight = FullPrnHeight * pic.Height / picPaper.Height
   ' Scale width & height
   PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
   PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
   
   ' Adjust left & top
   PrnLeft = FullPrnWidth * (pic.Left / picPaper.Width)
   PrnTop = FullPrnHeight * (pic.Top / picPaper.Height)
   ' Scale left & top
   PrnPicLeft = Prn.ScaleX(PrnLeft, vbHiMetric, Prn.ScaleMode)
   PrnPicTop = Prn.ScaleX(PrnTop, vbHiMetric, Prn.ScaleMode)
   
   Prn.PaintPicture picIN.Image, PrnPicLeft, PrnPicTop, PrnPicWidth, PrnPicHeight
''  or
'   Printer.ScaleMode = picIN.ScaleMode
'   SetStretchBltMode Printer.hdc, &H4 'HALFTONE
'   Call StretchDIBits(Printer.hdc, PrnPicLeft, PrnPicTop, PrnPicWidth, PrnPicHeight, 0, 0, picIN.Width, picIN.Height, iDATA(0, 0, 0), DIBInfo.bmiHeader, 0, vbSrcCopy)

End Sub

Private Sub optPaperSize_Click(Index As Integer)
' 0  A4   210x297  Portrait  297x210 Landscape
' 1  USA  216x279  Portrait  279x216 Landscape
   
   If Not aOpt Then Exit Sub
   
   PaperType = Index
   ' Set default Portrait
   If Index = 0 Then    ' A4
         paperW = 210
         paperH = 297
   Else                  ' USA
         paperW = 216
         paperH = 279
   End If
   
   If gPrintetOrientation = vbPRORPortrait Then
      optOrient_Click 0
   Else
      optOrient_Click 1
   End If
  
End Sub

Private Sub optOrient_Click(Index As Integer)
' 0  Portrait  PaperType = 0  A4 210x297, PaperType = 1 USA 8.5x11" 216x279
' 1  LandScape PaperType = 0  A4 297x210, PaperType = 1 USA 8.5x11" 270x216

Dim zAspect As Single
Dim picPaperW As Long, picPaperH As Long
Dim picW As Long, picH As Long
Dim PW$, PH$

   If Not aOpt Then Exit Sub
  
    ' Pre-calculate to see if it fits
   If Index = 0 Then   ' Portrait
         picPaperW = paperW
         picPaperH = paperH
         imgPrinterOrien.Picture = imgPage(Index).Picture
         gPrintetOrientation = vbPRORPortrait
         With picPaper
           .Width = picPaperW
           .Left = Shape1.Left + (Shape1.Width - paperW) / 2
           .Height = picPaperH
           .Top = Shape1.Top + (Shape1.Height - paperH) / 2
         End With
         
      
   Else  ' LandScape
         picPaperW = paperH
         picPaperH = paperW
         imgPrinterOrien.Picture = imgPage(Index).Picture
         gPrintetOrientation = vbPRORLandscape
         With picPaper
            .Width = picPaperW
            .Left = Shape1.Left + (Shape1.Width - paperH) / 2
            .Top = Shape1.Top + (Shape1.Height - paperW) / 2
            .Height = picPaperH
         End With
   End If
   
   zAspect = FullWidth / FullHeight
   If FullWidth > picPaperW Or FullHeight > picPaperH Then
      If FullWidth > FullHeight Then ' Landscape
         If FullWidth > picPaperW Then
            picW = picPaperW
            picH = picPaperW / zAspect
         Else
            picH = picPaperH
            picW = picPaperH * zAspect
         End If
      ElseIf FullWidth < FullHeight Then    ' Portrait
         If FullHeight > picPaperH Then
            picH = picPaperH
            picW = picPaperH * zAspect
         End If
      Else  ' FullWidth=FullHeight
         If Index = 0 Then ' Portrait
            picW = picPaperW
            picH = picPaperW
         Else   ' Landscape
            picW = picPaperH
            picH = picPaperH
         End If
      End If
   Else  ' smaller than or same size as paper
      If Index = 0 Then
         picW = FullWidth
         picH = FullHeight
      Else  ' Index = 1 Landscape
         picW = FullWidth
         picH = FullHeight
      End If
   End If
   
   If Multiplier > 1 Then
      If picW * 1 * Multiplier > picPaper.Width Or picH * 1 * Multiplier > picPaper.Height Then
         Multiplier = 1
         MsgBox "Too big - reduce scaling   ", vbInformation, "Printing"
         scrPercent.Value = 100
         Exit Sub
      End If
   End If
   
   InchMetricGrid
   
   ' Fits - proceed
   pic.Width = picW * 1 * Multiplier
   pic.Height = picH * 1 * Multiplier
   pic.Left = (picPaper.Width - pic.Width) / 2 - 1
   pic.Top = (picPaper.Height - pic.Height) / 2 - 1
   
   PicRefresh pic
   pic.Refresh
   
   ' Approximate print size width x height cm or inches
   PW$ = Format$(Round(pic.Width / StepW, 1), "#0.0")
   PH$ = Format$(Round(pic.Height / StepH, 1), "#0.0")
   If PaperType = 0 Then   ' A4
      LabWH = PW$ & " by " & PH$ & " cm"
   Else  ' US
      LabWH = PW$ & " by " & PH$ & " """
   End If
End Sub

Private Sub InchMetricGrid()
' A4 8.26" x 11.69"  PaperType = 0
' US 8.5"  x 11"
Dim k As Single
Dim n As Long
   Cls
   
   If gPrintetOrientation = vbPRORPortrait Then
      ' A4 8.26" x 11.69"   21 x 29.7 cm
      ' US 8.5"  x 11"
      If PaperType = 0 Then   ' A4
         StepW = picPaper.Width / 21
         StepH = picPaper.Height / 29.7
      Else  ' US
         StepW = picPaper.Width / 8.5
         StepH = picPaper.Height / 11
      End If
   
   Else  ' Landscape
      ' A4 11.69" x 8.26"   29.7 x 21 cm
      ' US 11"    x 8.5"
      If PaperType = 0 Then   ' A4
         StepW = picPaper.Width / 29.7
         StepH = picPaper.Height / 21
      Else  ' US
         StepW = picPaper.Width / 11
         StepH = picPaper.Height / 8.5
      End If
   
   End If
   n = 0
   For k = 0 To picPaper.Height Step StepH
      If n Mod 5 = 0 Then
         Me.Line (picPaper.Left - 10, picPaper.Top + k)-(picPaper.Left, picPaper.Top + k), vbRed, BF
      Else
         Me.Line (picPaper.Left - 6, picPaper.Top + k)-(picPaper.Left, picPaper.Top + k), vbRed, BF
      End If
      n = n + 1
   Next k
   n = 0
   For k = 0 To picPaper.Width Step StepW
      If n Mod 5 = 0 Then
         Me.Line (picPaper.Left + k, picPaper.Top)-(picPaper.Left + k, picPaper.Top - 10), vbRed, BF
      Else
         Me.Line (picPaper.Left + k, picPaper.Top)-(picPaper.Left + k, picPaper.Top - 6), vbRed, BF
      End If
      n = n + 1
   Next k
End Sub



Private Sub scrPercent_Change()
   Call scrPercent_Scroll
End Sub

Private Sub scrPercent_Scroll()
   Percent = scrPercent.Value
   LabPercent = Str$(Percent) & " %"
   Multiplier = Percent / 100
   
   If gPrintetOrientation = vbPRORPortrait Then
      optOrient_Click 0
   Else
      optOrient_Click 1
   End If
   
End Sub


Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   aMouseDown = True
   xp1 = X
   yp1 = Y
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NewX As Long, NewY As Long

   If aMouseDown Then
      NewX = pic.Left + (X - xp1)
      NewY = pic.Top + (Y - yp1)
       If NewX <= 5 Then NewX = 5
       If NewY <= 5 Then NewY = 5
       If NewX + pic.Width < picPaper.Width - 5 Then
       If NewY + pic.Height < picPaper.Height - 5 Then
            pic.Left = NewX
            pic.Top = NewY
       End If
       End If
   End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   aMouseDown = False
End Sub


Private Sub cmdExit_Click()
    Canceled = True
    Hide
End Sub

' Display the form. Return True if the user cancels.
Public Function ShowForm(ByVal cPic As Picture) As Boolean
   Dim PLIVE As Long
   Dim Prt As SelectPrinter
      
   Set picIN.Picture = cPic
   
   FullWidth = picIN.Width
   FullHeight = picIN.Height
   
   cmdOpen
     
   Dim FW As Long, FH As Long
   
   ' Set default
   Multiplier = 1
   
   aOpt = False
   ' Public FullWidth & FullHeight from main program
   optPaperSize(0).Value = True  ' A4
   If FullHeight >= FullWidth Then
      gPrintetOrientation = vbPRORPortrait    ' Taller than wide.
      optOrient(0).Value = True
   Else
      gPrintetOrientation = vbPRORLandscape   ' Wider than tall.
      optOrient(1).Value = True
   End If
   aOpt = True
   
   optPaperSize_Click 0
   
   InchMetricGrid
   
   Percent = 100
   scrPercent.Value = Percent
   LabPercent = Str$(Percent) & " %"
   
   'Refresh
   Show vbModal
      
   ShowForm = Canceled
   If Canceled = False Then
      picIN.Width = FullWidth
      picIN.Height = FullHeight
      picIN.Refresh
      PicRefresh picIN
   
      ' Calls Printer Common Dialog Box and sets
      ' the new printer temporarily for this print job.
      Printer.Orientation = gPrintetOrientation
   
      Prt = ShowPrinter(Me.hWnd, False)
      
      Refresh
   
      If Prt.bCanceled = False Then
         Printer.Orientation = gPrintetOrientation
         Printer.Print " ";
         PrintPicture Printer, picPaper, pic, picIN
         Printer.EndDoc
      End If
   End If
   
   Set picIN.Picture = LoadPicture()
   Set cPic = Nothing
   Erase iDATA
   Unload Me
End Function

Private Sub PicRefresh(pic As PictureBox)
  
  SetStretchBltMode pic.hDC, &H4 'HALFTONE
  Call StretchDIBits(pic.hDC, 0, 0, pic.Width, pic.Height, 0, 0, FullWidth, FullHeight, iDATA(0, 0, 0), DIBInfo.bmiHeader, 0, vbSrcCopy)
 
End Sub

Private Sub cmdOpen()
   
  Dim ratio As Single
        
  Dim hdcNew As Long
  Dim oldhand As Long
  Dim BytesPerScanLine As Long
  Dim PadBytesPerScanLine As Long
        
  'On Error Resume Next
  
  hdcNew = CreateCompatibleDC(GetDC(0))
  With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = picIN.Width
    .biHeight = picIN.Height     'bottom up scan line is now inverted
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = 0
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    PadBytesPerScanLine = BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With
  ReDim iDATA(0 To 3, 0 To picIN.Width - 1, 0 To picIN.Height - 1) As Byte
  'get bytes
   GetDIBits hdcNew, picIN, 0, picIN.Height, iDATA(0, 0, 0), DIBInfo, 0&
  DeleteDC hdcNew

End Sub
