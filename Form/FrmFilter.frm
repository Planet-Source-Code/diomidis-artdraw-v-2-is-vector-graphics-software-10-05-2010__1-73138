VERSION 5.00
Begin VB.Form FrmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter effect"
   ClientHeight    =   4530
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PictureOrig 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   990
      Left            =   2160
      ScaleHeight     =   930
      ScaleWidth      =   1695
      TabIndex        =   23
      Top             =   5310
      Width           =   1755
   End
   Begin VB.CommandButton CommandEffect 
      Caption         =   "Apply Effect"
      Height          =   450
      Left            =   9285
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Apply Effect in image"
      Top             =   3060
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton ComCancelFilter 
      Caption         =   "Cancel"
      Height          =   450
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Cancel Effect image"
      Top             =   3060
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox TmpPic 
      AutoSize        =   -1  'True
      Height          =   615
      Left            =   405
      ScaleHeight     =   555
      ScaleWidth      =   780
      TabIndex        =   21
      Top             =   5310
      Width           =   840
   End
   Begin VB.CommandButton CmdZoom 
      Height          =   465
      Index           =   2
      Left            =   4020
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FrmFilter.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   360
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton CmdZoom 
      Caption         =   "-"
      Height          =   315
      Index           =   1
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   810
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton CmdZoom 
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   4035
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   810
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4050
      Picture         =   "FrmFilter.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Copy in Original"
      Top             =   1830
      Width           =   600
   End
   Begin VB.CommandButton ComCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3855
      Width           =   1215
   End
   Begin VB.CommandButton ComOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4035
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3855
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3420
      Left            =   8685
      TabIndex        =   6
      Top             =   240
      Width           =   2415
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         LargeChange     =   10
         Left            =   180
         SmallChange     =   10
         TabIndex        =   14
         Top             =   2370
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.CommandButton Command5 
         Height          =   510
         Left            =   1725
         Picture         =   "FrmFilter.frx":06CC
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Select color"
         Top             =   1560
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   510
         Left            =   135
         ScaleHeight     =   450
         ScaleWidth      =   1500
         TabIndex        =   12
         Top             =   1590
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Filters"
         Top             =   900
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Group filters"
         Top             =   345
         Width           =   2190
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Color"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1410
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Filter"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   10
         Top             =   705
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Group Filter"
         Height          =   195
         Left            =   105
         TabIndex        =   9
         Top             =   135
         Width           =   810
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   135
      Top             =   300
   End
   Begin ArtDraw.ScrolledPicture ScrolledWindow2 
      Height          =   3390
      Left            =   4695
      TabIndex        =   20
      Top             =   315
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   5980
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ArtDraw.ScrolledPicture ScrolledWindow1 
      Height          =   3405
      Left            =   120
      TabIndex        =   19
      Top             =   300
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   6006
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Image with effect applied"
      Height          =   210
      Index           =   1
      Left            =   4650
      TabIndex        =   4
      Top             =   60
      Width           =   3765
   End
   Begin VB.Label Label1 
      Caption         =   "Original"
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   45
      Width           =   3765
   End
   Begin VB.Menu mnuDefinition 
      Caption         =   "Definition"
      Visible         =   0   'False
      Begin VB.Menu mnuBlur 
         Caption         =   "Smooth"
      End
      Begin VB.Menu mnuBlurMore 
         Caption         =   "Blur"
      End
      Begin VB.Menu mnuSharpen 
         Caption         =   "Sharpen"
      End
      Begin VB.Menu mnuSharpenMore 
         Caption         =   "Sharpen More"
      End
      Begin VB.Menu mnuDiffuse 
         Caption         =   "Diffuse"
      End
      Begin VB.Menu mnuDiffuseMore 
         Caption         =   "Diffuse More"
      End
      Begin VB.Menu mnuPixelize 
         Caption         =   "Pixelize"
      End
   End
   Begin VB.Menu mnuEdges 
      Caption         =   "Edges"
      Visible         =   0   'False
      Begin VB.Menu mnuEmboss 
         Caption         =   "Emboss"
      End
      Begin VB.Menu mnuEmbossMore 
         Caption         =   "Emboss More"
      End
      Begin VB.Menu mnuEngrave 
         Caption         =   "Engrave"
      End
      Begin VB.Menu mnuEngraveMore 
         Caption         =   "Engrave More"
      End
      Begin VB.Menu mnuRelief 
         Caption         =   "Relief"
      End
      Begin VB.Menu mnuEdge 
         Caption         =   "Edge Enhance"
      End
      Begin VB.Menu mnuContour 
         Caption         =   "Contour"
      End
      Begin VB.Menu mnuConnection 
         Caption         =   "Connected Contour"
      End
   End
   Begin VB.Menu mnuColors 
      Caption         =   "Colors"
      Visible         =   0   'False
      Begin VB.Menu mnuGreyScale 
         Caption         =   "GreyScale"
      End
      Begin VB.Menu mnuBlackWhite 
         Caption         =   "Black && White"
         Begin VB.Menu mnuBW1 
            Caption         =   "Nearest Color"
         End
         Begin VB.Menu mnuBW2 
            Caption         =   "Enhanced Diffusion"
         End
         Begin VB.Menu mnuBW3 
            Caption         =   "Ordered Dither"
         End
         Begin VB.Menu mnuBW4 
            Caption         =   "Floyd-Steinberg"
         End
         Begin VB.Menu mnuBW5 
            Caption         =   "Burke"
         End
         Begin VB.Menu mnuBW6 
            Caption         =   "Stucki"
         End
      End
      Begin VB.Menu mnuNegative 
         Caption         =   "Negative"
      End
      Begin VB.Menu mnuSwapColors 
         Caption         =   "Swap Colors"
         Begin VB.Menu mnuSwapBank 
            Caption         =   "RGB -> BRG"
            Index           =   1
         End
         Begin VB.Menu mnuSwapBank 
            Caption         =   "RGB -> GBR"
            Index           =   2
         End
         Begin VB.Menu mnuSwapBank 
            Caption         =   "RGB -> RBG"
            Index           =   3
         End
         Begin VB.Menu mnuSwapBank 
            Caption         =   "RGB -> BGR"
            Index           =   4
         End
         Begin VB.Menu mnuSwapBank 
            Caption         =   "RGB -> GRB"
            Index           =   5
         End
      End
      Begin VB.Menu mnuAqua 
         Caption         =   "Aqua"
      End
      Begin VB.Menu mnuAddNoise 
         Caption         =   "Add Noise"
      End
      Begin VB.Menu mnuGamma 
         Caption         =   "Gamma Correction"
      End
   End
   Begin VB.Menu mnuIntensity 
      Caption         =   "Intensity"
      Visible         =   0   'False
      Begin VB.Menu mnuBrighter 
         Caption         =   "Brighter"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuDarker 
         Caption         =   "Darker"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuContrast1 
         Caption         =   "Increase Contrast"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuContrast2 
         Caption         =   "Decrease Contrast"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuDilate 
         Caption         =   "Dilate"
      End
      Begin VB.Menu mnuErode 
         Caption         =   "Erode"
      End
      Begin VB.Menu mnuStretch 
         Caption         =   "Contrast Stretch"
      End
      Begin VB.Menu mnuSaturationI 
         Caption         =   "Increase Saturation"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuSaturationD 
         Caption         =   "Decrease Saturation"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "FrmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Canceled As Boolean
Private pProgress As Long
'Private TmpPic As StdPicture

' Display the form. Return True if the user cancels.
Public Function ShowForm(m_PicName As StdPicture) As Boolean  'm_PicName As String) As Boolean
    ' Assume we will cancel.
    Canceled = True
    Set TmpPic.Picture = m_PicName
   ' Set PictureEffect.Picture = TmpPic
    Set PictureOrig.Picture = TmpPic
      PictureOrig.Refresh
    Set ScrolledWindow1.Picture = TmpPic
    Set ScrolledWindow2.Picture = TmpPic
   ' PictureEffect.Move 0, 0
   ' PictureEffect.Refresh
   ' PictureOrig.Move 0, 0
  '  PictureOrig.Refresh
    ScrolledWindow1.Move 120, 330, 3900, 3400
    ScrolledWindow2.Move 4700, 330, 3900, 3400
    ScrolledWindow1.Refresh
    ScrolledWindow2.Refresh
    ComCancelFilter.Move CommandEffect.Left, CommandEffect.Top, CommandEffect.Width, CommandEffect.Height
    ' Display the form.
    Show vbModal

    ShowForm = Canceled
    If Not Canceled Then
        On Error Resume Next
        Set m_PicName = PictureOrig.Picture
        On Error GoTo 0
    Else
      Set m_PicName = Nothing
    End If
    Unload Me
End Function

Private Sub RunFilter(Filter As iFilterG, Factor As Long)
      Dim z As Double, VS As Long, HS As Long
        Timer1.Enabled = True
        Screen.MousePointer = 11
        'Set PictureEffect.Picture = PictureOrig.Image
        Set ScrolledWindow2.Picture = PictureOrig.Image
        TmpPic.Picture = PictureOrig.Image
        TmpPic.Refresh
       
        ComCancelFilter.Visible = True
        ComOK.Enabled = False
        ComCancel.Enabled = False
        Frame1.Enabled = False
        
        ComCancelFilter.ZOrder
        If FilterG(Filter, TmpPic.Image, Factor, pProgress) = False Then
           'Set PictureEffect.Picture = TmpPic.Image
           Set ScrolledWindow2.Picture = TmpPic.Image
        End If
        ComCancelFilter.Visible = False
        ComOK.Enabled = True
        ComCancel.Enabled = True
        Frame1.Enabled = True
        Screen.MousePointer = 0
End Sub



Private Sub Combo1_Click()
     If Combo1.ListIndex = -1 Then Exit Sub
     'Set PictureEffect.Picture = PictureOrig.Image
     Set ScrolledWindow2.Picture = PictureOrig.Image
     Combo2.Clear
     Command5.Visible = False
     Picture1.Visible = False
     HScroll1.Visible = False
     HScroll1.SmallChange = 1
     HScroll1.LargeChange = 10
     CommandEffect.Visible = False
     ComCancelFilter.Visible = False
     Label4(0).Visible = False
     Label3.Caption = ""
     
     Select Case Combo1.ItemData(Combo1.ListIndex)
     Case 1 'Definition
            Combo2.AddItem "Smooth": Combo2.ItemData(Combo2.NewIndex) = 5          '------
            Combo2.AddItem "Blur": Combo2.ItemData(Combo2.NewIndex) = 3            '------
            Combo2.AddItem "Sharpen": Combo2.ItemData(Combo2.NewIndex) = 1         '2  '0-N +-
            Combo2.AddItem "Sharpen More": Combo2.ItemData(Combo2.NewIndex) = 101  '0  '0-N +-
            Combo2.AddItem "Diffuse": Combo2.ItemData(Combo2.NewIndex) = 4         '6
            Combo2.AddItem "Diffuse More": Combo2.ItemData(Combo2.NewIndex) = 104  '12
            Combo2.AddItem "Pixelize": Combo2.ItemData(Combo2.NewIndex) = 15       '--size Pix
            Combo2.AddItem "Rect": Combo2.ItemData(Combo2.NewIndex) = 45       '--
            Combo2.AddItem "Fog": Combo2.ItemData(Combo2.NewIndex) = 40
     Case 2 'Edges
            Combo2.AddItem "Emboss": Combo2.ItemData(Combo2.NewIndex) = 8                    '--RGB----
            Combo2.AddItem "Emboss More": Combo2.ItemData(Combo2.NewIndex) = 9               '--RGB----
            Combo2.AddItem "Engrave": Combo2.ItemData(Combo2.NewIndex) = 10                  '--RGB----
            Combo2.AddItem "Engrave More": Combo2.ItemData(Combo2.NewIndex) = 11             '--RGB----
            Combo2.AddItem "Relief": Combo2.ItemData(Combo2.NewIndex) = 13                   '---------
            Combo2.AddItem "Edge Enhance": Combo2.ItemData(Combo2.NewIndex) = 6              '0-N +-
            Combo2.AddItem "Contour": Combo2.ItemData(Combo2.NewIndex) = 7                   '--RGB----
            Combo2.AddItem "Connected Contour": Combo2.ItemData(Combo2.NewIndex) = 27        '-------
            Combo2.AddItem "Neon": Combo2.ItemData(Combo2.NewIndex) = 32                     '-------
            Combo2.AddItem "Art": Combo2.ItemData(Combo2.NewIndex) = 38
            Combo2.AddItem "Snow": Combo2.ItemData(Combo2.NewIndex) = 41
            Combo2.AddItem "Wave": Combo2.ItemData(Combo2.NewIndex) = 42
            Combo2.AddItem "Crease": Combo2.ItemData(Combo2.NewIndex) = 43
            Combo2.AddItem "Stranges": Combo2.ItemData(Combo2.NewIndex) = 39

      Case 3 'Colors
            Combo2.AddItem "Grey Scale": Combo2.ItemData(Combo2.NewIndex) = 12              '------
            Combo2.AddItem "Negative": Combo2.ItemData(Combo2.NewIndex) = 2                 '------
            Combo2.AddItem "Aqua": Combo2.ItemData(Combo2.NewIndex) = 24                    '-----
            Combo2.AddItem "Add Noise": Combo2.ItemData(Combo2.NewIndex) = 29               '0
            Combo2.AddItem "Gamma Correction": Combo2.ItemData(Combo2.NewIndex) = 31        '1-100
            Combo2.AddItem "Sepia": Combo2.ItemData(Combo2.NewIndex) = 44
            Combo2.AddItem "Ice": Combo2.ItemData(Combo2.NewIndex) = 47
            Combo2.AddItem "Comic": Combo2.ItemData(Combo2.NewIndex) = 46
            'Black & White
            Combo2.AddItem "B&W Nearest Color": Combo2.ItemData(Combo2.NewIndex) = 18       '--RGB--
            Combo2.AddItem "B&W Enhanced Diffusion": Combo2.ItemData(Combo2.NewIndex) = 19  '-----
            Combo2.AddItem "B&W Ordered Dither": Combo2.ItemData(Combo2.NewIndex) = 20      '-----
            Combo2.AddItem "B&W Floyd-Steinberg": Combo2.ItemData(Combo2.NewIndex) = 21    '1-n Palette
            Combo2.AddItem "B&W Burke": Combo2.ItemData(Combo2.NewIndex) = 22               '1-n Palette
            Combo2.AddItem "B&W Stucki": Combo2.ItemData(Combo2.NewIndex) = 23              '1-n Palette
            'Swap Colors
            Combo2.AddItem "Swap Colors -> BRG": Combo2.ItemData(Combo2.NewIndex) = 161
            Combo2.AddItem "Swap Colors -> GBR": Combo2.ItemData(Combo2.NewIndex) = 162
            Combo2.AddItem "Swap Colors -> RBG": Combo2.ItemData(Combo2.NewIndex) = 163
            Combo2.AddItem "Swap Colors -> BGR": Combo2.ItemData(Combo2.NewIndex) = 164
            Combo2.AddItem "Swap Colors -> GRB": Combo2.ItemData(Combo2.NewIndex) = 165
       Case 4 'Intensity
            Combo2.AddItem "Brighter": Combo2.ItemData(Combo2.NewIndex) = 14 '- (10)          '>0
            Combo2.AddItem "Darker": Combo2.ItemData(Combo2.NewIndex) = 141 ' - (-10)           '<0
            Combo2.AddItem "Increase Contrast": Combo2.ItemData(Combo2.NewIndex) = 17        '>0
            Combo2.AddItem "Decrease Contrast": Combo2.ItemData(Combo2.NewIndex) = 171        '<0
            Combo2.AddItem "Dilate": Combo2.ItemData(Combo2.NewIndex) = 25                   '-------
            Combo2.AddItem "Erode": Combo2.ItemData(Combo2.NewIndex) = 26                    '-------
            Combo2.AddItem "Contrast Stretch": Combo2.ItemData(Combo2.NewIndex) = 28         '------
            Combo2.AddItem "Increase Saturation": Combo2.ItemData(Combo2.NewIndex) = 30      '>0  15
            Combo2.AddItem "Decrease Saturation": Combo2.ItemData(Combo2.NewIndex) = 301      '<0  (-20)
       Case 5
            Combo2.AddItem "3d Grid": Combo2.ItemData(Combo2.NewIndex) = 33
            Combo2.AddItem "MirrorRL": Combo2.ItemData(Combo2.NewIndex) = 34
            Combo2.AddItem "MirrorLR": Combo2.ItemData(Combo2.NewIndex) = 35
            Combo2.AddItem "MirrorDT": Combo2.ItemData(Combo2.NewIndex) = 36
            Combo2.AddItem "MirrorTD": Combo2.ItemData(Combo2.NewIndex) = 37
       End Select
       
End Sub

Private Sub Combo2_Click()
     '
     If Combo1.ListIndex = -1 Then Exit Sub
     If Combo2.ListIndex = -1 Then Exit Sub
     'Set PictureEffect.Picture = PictureOrig.Image
     Set ScrolledWindow2.Picture = PictureOrig.Image
     
     Command5.Visible = False
     Picture1.Visible = False
     HScroll1.Visible = False
     HScroll1.SmallChange = 1
     HScroll1.LargeChange = 10
     CommandEffect.Visible = False
     ComCancelFilter.Visible = False
     Label4(0).Visible = False
     Label3.Caption = ""
              
     Select Case Combo1.ItemData(Combo1.ListIndex)
     Case 1 'Definition
         Select Case Combo2.ItemData(Combo2.ListIndex)
         Case 5 ' Smooth-5          '------
               RunFilter iSmooth, 0
         Case 3 ' Blur-3            '------
              RunFilter iBlur, 0
         Case 1 'Sharpen-1         '2  '0-N +-
              RunFilter iSharpen, 2
         Case 101 'Sharpen More- 101  '0  '0-N +-
              RunFilter iSharpen, 0
         Case 4 'Diffuse-4         '6
              RunFilter iDiffuse, 6
         Case 104 'Diffuse More- 104  '12
              RunFilter iDiffuse, 12
         Case 15 'Pixelize-15       '--size Pix
              HScroll1.Visible = True
              CommandEffect.Visible = True
              HScroll1.Max = 100
              HScroll1.Min = 1
              HScroll1.Value = 5
              Label3.Caption = HScroll1.Value
         Case 45 'Rect
              RunFilter iRects, 12
         Case 40
              HScroll1.Visible = True
              CommandEffect.Visible = True
              HScroll1.Max = 100
              HScroll1.Min = 1
              HScroll1.Value = 50
              Label3.Caption = HScroll1.Value
         End Select
     Case 2 'Edges
         Select Case Combo2.ItemData(Combo2.ListIndex)
         Case 7, 8, 9, 10, 11:
               'Emboss-8                    '--RGB----
               'Emboss More-9               '--RGB----
               'Engrave-10                  '--RGB----
               'Engrave More-11             '--RGB----
                'Contour-7                   '--RGB----
               CommandEffect.Visible = True
               Label4(0).Visible = True
               Command5.Visible = True
               Picture1.Visible = True
                              
         Case 13: 'Relief-13                   '---------
               RunFilter iRelief, 0
         Case 6: 'Edge Enhance-6              '1-N +-
              HScroll1.Visible = True
              CommandEffect.Visible = True
              HScroll1.SmallChange = 1
              HScroll1.LargeChange = 2
              HScroll1.Max = 10
              HScroll1.Min = 1
              HScroll1.Value = 1
              Label3.Caption = HScroll1.Value
         Case 27: 'Connected Contour-27        '-------
              RunFilter iConnection, 0
         Case 32: 'Neon-32                     '-------
              RunFilter iNeon, 0
         
         Case 38 'Art:
              RunFilter iArt, 0
         
         Case 39, 41 'Stranges, Snow
              HScroll1.Visible = True
              CommandEffect.Visible = True
              HScroll1.Max = 100
              HScroll1.Min = 1
              HScroll1.Value = 50
              Label3.Caption = HScroll1.Value
             
         Case 42 'Wave
              HScroll1.Visible = True
              CommandEffect.Visible = True
              HScroll1.SmallChange = 1
              HScroll1.LargeChange = 4
              HScroll1.Max = 16
              HScroll1.Min = 0
              HScroll1.Value = 5
              Label3.Caption = HScroll1.Value
          Case 43 'Crease
              HScroll1.Visible = True
              CommandEffect.Visible = True
              HScroll1.SmallChange = 10
              HScroll1.LargeChange = 100
              HScroll1.Max = 1024
              HScroll1.Min = 64
              HScroll1.Value = 512
              Label3.Caption = HScroll1.Value
         End Select
         
      Case 3 'Colors
          Select Case Combo2.ItemData(Combo2.ListIndex)
          Case 12: 'Grey Scale-12              '------
                RunFilter iGreyScale, 0
          Case 2: 'Negative-2                 '------
                RunFilter iNegative, 0
          Case 24: 'Aqua-24                    '-----
                RunFilter iAqua, 0
          Case 29: 'Add Noise-29               '0
                HScroll1.Visible = True
                CommandEffect.Visible = True
                HScroll1.SmallChange = 10
                HScroll1.LargeChange = 100
                HScroll1.Max = 500
                HScroll1.Min = 0
                HScroll1.Value = 50
               Label3.Caption = HScroll1.Value
          Case 31: 'Gamma Correction-31        '1-100
                HScroll1.Visible = True
                CommandEffect.Visible = True
                HScroll1.Max = 100
                HScroll1.Min = 1
                HScroll1.Value = 50
                Label3.Caption = HScroll1.Value
          'Black & White
          Case 18: 'B&W Nearest Color-18       '--RGB--
               RunFilter iColDepth1, RGB(180, 180, 180)
          Case 19: '"B&W Enhanced Diffusion-19  '-----
               RunFilter iColDepth2, 0
          Case 20: 'B&W Ordered Dither-20      '-----
               RunFilter iColDepth3, 0
          Case 21: 'B&W Floyd-Steinberg-21    '1-n Palette
               RunFilter iColDepth4, 15
          Case 22: 'B&W Burke-22              '1-n Palette
               RunFilter iColDepth5, 15
          Case 23: 'B&W Stucki-23             '1-n Palette
               RunFilter iColDepth6, 15
          'Swap Colors
          Case 161: 'Swap Colors -> BRG-161
               RunFilter iSwapBank, 1
          Case 162: 'Swap Colors -> GBR-162
               RunFilter iSwapBank, 2
          Case 163: 'Swap Colors -> RBG-163
               RunFilter iSwapBank, 3
          Case 164: 'Swap Colors -> BGR-164
               RunFilter iSwapBank, 4
          Case 165: 'Swap Colors -> GRB-165
               RunFilter iSwapBank, 5
          Case iSepia:
               RunFilter iSepia, 0
          Case 46 ' Comic
               RunFilter iComic, 0
          Case 47 'Ice
               RunFilter iIce, 0
          End Select
       Case 4 'Intensity
           Select Case Combo2.ItemData(Combo2.ListIndex)
           Case 25: 'Dilate-25                  '------
                RunFilter iDilate, 0
           Case 26: ' Erode-26                   '------
                RunFilter iErode, 0
           Case 28: ' Contrast Stretch-28        '------
                RunFilter iStretch, 0
           Case Else
            'Brighter-14              '>0
            'Darker-141               '<0
            'Increase Contrast-17     '>0
            'Decrease Contrast-171    '<0
            'Increase Saturation-30   '>0  15
            'Decrease Saturation-301  '<0  (-20)
                HScroll1.Visible = True
                CommandEffect.Visible = True
                HScroll1.Max = 100
                HScroll1.Min = 1
                If Combo2.ItemData(Combo2.ListIndex) = 30 Then
                  HScroll1.Value = 15
                ElseIf Combo2.ItemData(Combo2.ListIndex) = 301 Then
                  HScroll1.Value = 20
                Else
                  HScroll1.Value = 10
                End If
                Label3.Caption = HScroll1.Value
           End Select
       Case 5
         Select Case Combo2.ItemData(Combo2.ListIndex)
         Case iMirrorRL: RunFilter iMirrorRL, 0
         Case iMirrorLR: RunFilter iMirrorLR, 0
         Case iMirrorDT: RunFilter iMirrorDT, 0
         Case iMirrorTD: RunFilter iMirrorTD, 0
         Case Else
             HScroll1.Visible = True
             CommandEffect.Visible = True
             HScroll1.Max = 100
             HScroll1.Min = 1
             HScroll1.Value = 50
             Label3.Caption = HScroll1.Value
         End Select
       End Select
End Sub

Private Sub ComCancelFilter_Click()
       mCancel = True
End Sub

Private Sub ComOK_Click()
        Canceled = False
        Me.Hide
End Sub

Private Sub ComCancel_Click()
       Canceled = True
       Hide
End Sub

Private Sub CommandEffect_Click()

 If Combo1.ListIndex = -1 Then Exit Sub
 If Combo2.ListIndex = -1 Then Exit Sub
 
     Select Case Combo1.ItemData(Combo1.ListIndex)
     Case 1 'Definition
         Select Case Combo2.ItemData(Combo2.ListIndex)
         Case 1 'Sharpen
              RunFilter iSharpen, HScroll1.Value
         Case 15 'Pixelize-15       '--size Pix
              RunFilter iPixelize, HScroll1.Value
         Case 40 'iFog
              RunFilter iFog, HScroll1.Value
         End Select
      Case 2
         Select Case Combo2.ItemData(Combo2.ListIndex)
         Case 8: 'Emboss--RGB----
              RunFilter iEmboss, Picture1.BackColor
         Case 9: 'Emboss More--RGB----
              RunFilter iEmbossMore, Picture1.BackColor
         Case 10: 'Engrave--RGB----
              RunFilter iEngrave, Picture1.BackColor
         Case 11: 'Engrave More'--RGB----
              RunFilter iEngraveMore, Picture1.BackColor
         Case 6: 'Edge Enhance1-N +-
              RunFilter iEDGE, HScroll1.Value
         Case 7: 'Contour--RGB----
              RunFilter iContour, Picture1.BackColor
          Case 39 'Stranges
              RunFilter iStranges, HScroll1.Value
          Case 41 'Snow:
              RunFilter iSnow, HScroll1.Value
          Case 42 'Wave:
              RunFilter iWave, HScroll1.Value '0-16
          Case 43 'Crease:
              RunFilter iCrease, HScroll1.Value '64-65536 >512
          End Select
      Case 3
           Select Case Combo2.ItemData(Combo2.ListIndex)
           Case 29:
                RunFilter iAddNoise, HScroll1.Value
           Case 31:
                RunFilter iGamma, HScroll1.Value
          
           Case Else
'              Stop
           End Select
      Case 4
           Select Case Combo2.ItemData(Combo2.ListIndex)
           Case 14:  'Brighter        '>0
                 RunFilter iBRIGHTNESS, HScroll1.Value
           Case 141: 'Darker"          '<0
                 RunFilter iBRIGHTNESS, -HScroll1.Value
           Case 17: 'Increase Contrast    '>0
                 RunFilter iContrast, HScroll1.Value
           Case 171: 'Decrease Contrast       '<0
                 RunFilter iContrast, -HScroll1.Value
           Case 30: 'Increase Saturation  '>0
                 RunFilter iSaturation, HScroll1.Value
           Case 301: 'Decrease Saturation '<0
                 RunFilter iSaturation, -HScroll1.Value
           End Select
      Case 5
            RunFilter iGrid3d, HScroll1.Value
      End Select
End Sub

Private Sub Command4_Click()
    
    Set ScrolledWindow1.Picture = TmpPic.Image
    Set PictureOrig.Picture = TmpPic.Image
    ScrolledWindow1.Refresh
    ScrolledWindow2.Refresh
End Sub

Private Sub Command5_Click()
    Dim NewColor As Long
    Screen.MousePointer = 11
    NewColor = Picture1.BackColor
    If frmVbDraw.ShowColor(NewColor) = True Then
       Picture1.BackColor = NewColor
    End If
    Screen.MousePointer = 0
End Sub


Private Sub Form_Load()
      Combo1.Clear
      Combo1.AddItem "Definition": Combo1.ItemData(Combo1.NewIndex) = 1
      Combo1.AddItem "Edges": Combo1.ItemData(Combo1.NewIndex) = 2
      Combo1.AddItem "Colors": Combo1.ItemData(Combo1.NewIndex) = 3
      Combo1.AddItem "Intensity": Combo1.ItemData(Combo1.NewIndex) = 4
      Combo1.AddItem "Effect": Combo1.ItemData(Combo1.NewIndex) = 5
      Combo1.ListIndex = 0
End Sub

Private Sub HScroll1_Change()
        Label3.Caption = HScroll1.Value
End Sub

Private Sub ScrolledWindow1_Scroll(HValue As Integer, Vvalue As Integer)
        ScrolledWindow2.Hscroll = HValue
        ScrolledWindow2.Vscroll = Vvalue
End Sub

Private Sub ScrolledWindow2_Scroll(HValue As Integer, Vvalue As Integer)
         ScrolledWindow1.Hscroll = HValue
         ScrolledWindow1.Vscroll = Vvalue
End Sub

Private Sub Timer1_Timer()
  Me.Caption = "Filter effect " & Str(pProgress) + "%"
  If pProgress = 100 Then
    Me.Caption = "Filter effect"
    Timer1.Enabled = False
  End If
End Sub


