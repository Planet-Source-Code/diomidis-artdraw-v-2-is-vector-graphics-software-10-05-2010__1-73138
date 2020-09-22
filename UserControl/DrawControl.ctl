VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl DrawControl 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9780
   KeyPreview      =   -1  'True
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   652
   Begin ArtDraw.MeForm MeForm4 
      Height          =   3855
      Left            =   4185
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   6800
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   75
         TabIndex        =   24
         Top             =   300
         Width           =   2205
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   2790
         Width           =   780
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Top             =   3105
         Width           =   780
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   315
         Left            =   1245
         TabIndex        =   21
         Top             =   3450
         Width           =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Id:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   28
         Top             =   2805
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1230
         TabIndex        =   27
         Top             =   2820
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1215
         TabIndex        =   26
         Top             =   3135
         Width           =   195
      End
      Begin VB.Label LabelId 
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   420
         TabIndex        =   25
         Top             =   2805
         Width           =   570
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1335
      LargeChange     =   50
      Left            =   6330
      Max             =   0
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3975
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   50
      Left            =   4815
      Max             =   0
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5535
      Visible         =   0   'False
      Width           =   1530
   End
   Begin ArtDraw.MeForm MeForm3 
      Height          =   2475
      Left            =   4575
      TabIndex        =   14
      Top             =   1500
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   4366
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin ArtDraw.CtrTranform CtrTranform1 
         Height          =   2130
         Left            =   105
         TabIndex        =   15
         Top             =   285
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   3757
      End
   End
   Begin ArtDraw.MeForm MeForm2 
      Height          =   3210
      Left            =   6705
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   5662
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton ComDropperFill 
         Height          =   315
         Left            =   270
         Picture         =   "DrawControl.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2295
         Width           =   345
      End
      Begin ArtDraw.CtlFill CtlFill1 
         Height          =   2790
         Left            =   15
         TabIndex        =   17
         Top             =   315
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   4921
         Color2          =   16777215
      End
   End
   Begin ArtDraw.MeForm MeForm1 
      Height          =   2850
      Left            =   6750
      TabIndex        =   3
      Top             =   705
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   5027
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ImageList imlDrawWidths 
         Left            =   1635
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":038A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":059C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":07AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":09C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":0BD2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":0DE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":0FF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":1208
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":141A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":162C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlDrawStyles 
         Left            =   1665
         Top             =   1410
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   40
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":183E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":1A10
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":1BE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":1DB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":1F86
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":2158
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageCombo icbDrawStyle 
         Height          =   330
         Left            =   90
         TabIndex        =   9
         Top             =   1935
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin MSComctlLib.ImageCombo icbDrawWidth 
         Height          =   330
         Left            =   105
         TabIndex        =   8
         Top             =   1290
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.CommandButton ComDropperPen 
         Height          =   315
         Left            =   120
         Picture         =   "DrawControl.ctx":232A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2340
         Width           =   360
      End
      Begin VB.CommandButton cmdSysColorsPen 
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   1650
         Picture         =   "DrawControl.ctx":26B4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " System colors "
         Top             =   555
         Width           =   375
      End
      Begin VB.CommandButton CommandPen 
         Caption         =   "Apply"
         Height          =   330
         Left            =   480
         TabIndex        =   5
         Top             =   2325
         Width           =   1485
      End
      Begin VB.PictureBox PicPenColor 
         BackColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   180
         ScaleHeight     =   405
         ScaleWidth      =   1380
         TabIndex        =   4
         Top             =   495
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pen Color "
         Height          =   180
         Left            =   195
         TabIndex        =   12
         Top             =   270
         Width           =   1800
      End
      Begin VB.Label LbDrawWidth 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Width"
         Height          =   180
         Left            =   165
         TabIndex        =   11
         Top             =   1035
         Width           =   1935
      End
      Begin VB.Label LblDrawStyle 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Style"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   1680
         Width           =   1815
      End
   End
   Begin VB.PictureBox picHidden 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   105
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   2
      Top             =   1245
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton ComCorner 
      Height          =   240
      Left            =   6405
      TabIndex        =   0
      Top             =   5535
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   840
      MousePointer    =   99  'Custom
      ScaleHeight     =   166
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   1
      Top             =   1365
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Line LineX 
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         Visible         =   0   'False
         X1              =   12
         X2              =   139
         Y1              =   20
         Y2              =   20
      End
      Begin VB.Line LineY 
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         Visible         =   0   'False
         X1              =   29
         X2              =   156
         Y1              =   44
         Y2              =   44
      End
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   22
      Left            =   4875
      Picture         =   "DrawControl.ctx":287E
      Top             =   690
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   21
      Left            =   3990
      Picture         =   "DrawControl.ctx":29D0
      Top             =   645
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   20
      Left            =   3435
      Picture         =   "DrawControl.ctx":2B22
      Top             =   675
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   19
      Left            =   2835
      Picture         =   "DrawControl.ctx":2C74
      Top             =   570
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   18
      Left            =   1605
      Picture         =   "DrawControl.ctx":2DC6
      Top             =   585
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   17
      Left            =   1110
      Picture         =   "DrawControl.ctx":2F18
      Top             =   540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   16
      Left            =   660
      Picture         =   "DrawControl.ctx":306A
      Top             =   585
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   15
      Left            =   420
      Picture         =   "DrawControl.ctx":31BC
      Top             =   615
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   14
      Left            =   2460
      Picture         =   "DrawControl.ctx":330E
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   13
      Left            =   2040
      Picture         =   "DrawControl.ctx":3460
      Top             =   690
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   12
      Left            =   1935
      Picture         =   "DrawControl.ctx":35B2
      Top             =   795
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   11
      Left            =   6705
      Picture         =   "DrawControl.ctx":3704
      Top             =   75
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   10
      Left            =   6015
      Picture         =   "DrawControl.ctx":3856
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   9
      Left            =   5460
      Picture         =   "DrawControl.ctx":3B60
      Top             =   105
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   8
      Left            =   4875
      Picture         =   "DrawControl.ctx":3E6A
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   7
      Left            =   4335
      Picture         =   "DrawControl.ctx":3FBC
      Top             =   105
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   6
      Left            =   3675
      Picture         =   "DrawControl.ctx":42C6
      Top             =   75
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   5
      Left            =   3060
      Picture         =   "DrawControl.ctx":45D0
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   4
      Left            =   2385
      Picture         =   "DrawControl.ctx":48DA
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   3
      Left            =   1725
      Picture         =   "DrawControl.ctx":4BE4
      Top             =   30
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   2
      Left            =   1035
      Picture         =   "DrawControl.ctx":4EEE
      Top             =   15
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   1
      Left            =   450
      Picture         =   "DrawControl.ctx":51F8
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   0
      Left            =   45
      Picture         =   "DrawControl.ctx":5AC2
      Top             =   540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuTransform 
      Caption         =   "Transform"
      Begin VB.Menu mnuClearTransform 
         Caption         =   "Clear Transform"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCurve 
         Caption         =   "Make Curve"
      End
      Begin VB.Menu mnuFillMode 
         Caption         =   "FillMode"
         Begin VB.Menu mnuAlternate 
            Caption         =   "Alternate"
         End
         Begin VB.Menu mnuWinding 
            Caption         =   "Winding"
         End
      End
      Begin VB.Menu seplock 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLock 
         Caption         =   "Lock"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUnlock 
         Caption         =   "UnLock"
      End
      Begin VB.Menu sepEditPoints 
         Caption         =   "-"
      End
      Begin VB.Menu mnueditPoints 
         Caption         =   "Edit points"
      End
      Begin VB.Menu sepProperty 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperty 
         Caption         =   "Property"
      End
   End
   Begin VB.Menu mnuEditPoint 
      Caption         =   "EditPoint"
      Begin VB.Menu mnuAddNode 
         Caption         =   "Add node(s)"
      End
      Begin VB.Menu mnuDeletenode 
         Caption         =   "Delete node(s)"
      End
      Begin VB.Menu sepDeleteNode 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToLine 
         Caption         =   "To Line"
      End
      Begin VB.Menu mnuToCurve 
         Caption         =   "To Curve"
      End
      Begin VB.Menu sepToCurve 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoClose 
         Caption         =   "Auto close"
      End
      Begin VB.Menu mnuautoopen 
         Caption         =   "Auto open"
      End
      Begin VB.Menu sepautoopen 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBreakNode 
         Caption         =   "Break node"
      End
      Begin VB.Menu mnuBreakApart 
         Caption         =   "Break Apart"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "DrawControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(c) 2007-2010 Diomidis Kiriakopoulos

Dim idredraw As Long

'Default Property Values:
Const m_def_ShowObjectPoint = 0
Const m_def_BackColor = &HFFFFFF
Const m_def_DrawRuler = 0
Const m_def_CrossMouse = 0
Const m_def_EditPoint = 0
Const m_def_hDC = 0
Const m_def_LockObject = 0
Const m_def_BackImage = ""
Const m_def_Blend = 0
Const m_def_FillColor2 = 0
Const m_def_Pattern = ""
Const m_def_TypeGradient = 0
Const m_def_ShowPenProperty = 0
Const m_def_ShowFillProperty = 0
Const m_def_m_ShowTranformProperty = False
Const m_def_FileTitle = ""
Const m_def_FileName = ""
Const m_def_ForeColor = 0
Const m_def_DrawWidth = 1
Const m_def_DrawStyle = 0
Const m_def_FillStyle = 0
Const m_def_FillColor = 0
Const m_def_ShowCanvasSize = False
Const m_def_CanvasWidth = 794
Const m_def_CanvasHeight = 1124

'Property Variables:
Dim m_ShowObjectPoint As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_DrawRuler As Boolean
Dim m_CrossMouse As Boolean
Dim m_EditPoint As Boolean
Dim m_ObjPicture As Picture
Dim m_hDC As Long
Dim m_LockObject As Boolean
Dim m_BackImage As String
Dim m_Blend As Integer
Dim m_FillColor2 As OLE_COLOR
Dim m_Pattern As String
Dim m_TypeGradient As Integer
Dim m_ShowPenProperty As Boolean
Dim m_ShowFillProperty As Boolean
Dim m_ShowTranformProperty As Boolean
Dim m_FileTitle As String
Dim m_FileName As String
Dim m_ForeColor As OLE_COLOR
Dim m_DrawWidth As Integer
Dim m_DrawStyle As Integer
Dim m_FillStyle As Integer
Dim m_FillColor As OLE_COLOR
Dim m_ShowCanvasSize As Boolean
Dim m_CanvasLeft As Long
Dim m_CanvasTop As Long
Dim m_CanvasWidth As Long
Dim m_CanvasHeight As Long
Dim m_Image As Picture
Dim mLockControl As Boolean

'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Event KeyPress(ByVal KeyAscii As Integer)
Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event SetDirty()
Event EnableMenusForSelection()
Event ColorSelected(ByVal tColor As Integer, ByVal cColor As Long)
Event MsgControl(ByVal txt As String)
Event EnableMenuText(ByVal MenuOn As Boolean)
Event EnableMenuBitMap(ByVal MenuOn As Boolean)
Event ZoomChange()
Event SizeCanvas(ByVal Width As Single, ByVal Height As Single)

Public Obj As vbdObject
Private MX() As Single, mY() As Single ', m_typePoint() As Byte, m_SelectPoint As Long, m_DrawingPoint As Boolean
Private e_NumPoints As Long, m_NumPoints As Long, m_TypePoint() As Byte, m_OriginalPoints() As PointAPI, m_SelectPoint As Long, m_DrawingPoint As Boolean

Private TmpObj As vbdObject
Private m_DrawingObject As Boolean
Private m_EditObject As Boolean

' Rubberband variables.
Private m_StartX As Single
Private m_StartY As Single
Private m_LastX As Single
Private m_LastY As Single
Private X1 As Single
Private X2 As Single
Private Y1 As Single
Private Y2 As Single

Public XminBox As Single
Public YminBox As Single
Public XmaxBox As Single
Public YmaxBox As Single

Private mPicture As Long 'StdPicture
Private Xmin_Box As Single
Private Ymin_Box As Single
Private Xmax_Box As Single
Private Ymax_Box As Single

'Private m_ReadFillProperty As Boolean
'Private m_ReadPenProperty As Boolean

Private m_Rotate As Boolean
Private m_Move As Boolean
Private m_Scale As Boolean
Private m_Skew As Boolean

Private m_ScaleType As Integer
Private Ortho As RectAngle

Dim m(1 To 3, 1 To 3) As Single
Dim MeFormView1 As Boolean
Dim MeFormView2 As Boolean
Dim MeFormView3 As Boolean
Dim MeFormView4 As Boolean

Public Enum m_Order
    BringToFront = 0
    SendToBack = 1
    BringFoward = 2
    SendBackward = 3
End Enum

' Global max and min world coordinates (including margins).
Private DataXmin As Long
Private DataXmax As Long
Private DataYmin As Long
Private DataYmax As Long

' Set the min and max allowed width and height.
Private DataMinWid As Long
Private DataMinHgt As Long
Private DataMaxWid As Long
Private DataMaxHgt As Long

' The aspect ratio of the viewport.
Private VAspect As Single
Private CenterZoomX As Long
Private CenterZoomY As Long
        
'' Current world window bounds.
''ÔñÝ÷ïõóá êüóìï ðáñÜèõñï üñéá.
'Private Wxmin As Single
'Private Wxmax As Single
'Private Wymin As Single
'Private Wymax As Single

' Prevent change events when we are adjusting the scroll bars.
Private IgnoreSbarChange As Boolean

' Variables used for zooming.
Public DrawingMode As Integer
Private OldDrawingMode As Integer
Const MODE_NONE = 0

Const MODE_EDITOBJ = 1
Const MODE_EDITPOINT = 2
Const MODE_POLYLINE = 3
Const MODE_FREELINE = 4
Const MODE_Scribble = 5
Const MODE_RECTANGLE = 8
Const MODE_POLYGON = 9
Const MODE_ELLIPSE = 10
Const MODE_TEXT = 11
Const MODE_TEXTFRAME = 12
Const MODE_PICTURE = 15
Const MODE_ZOOMING = 16
Const MODE_PANNING = 18
Const MODE_START_ZOOM = 19
Const MODE_ReadFill = 20
Const MODE_ReadPen = 21

Private ScrollMouse As Integer
Private ScrollType As Integer '

Private zStartX As Single
Private zStartY As Single
Private zLastX As Single
Private zLastY As Single
Private OldMode As Integer

Public Function AddPicture() As Long
    If mPicture = 0 Then mPicture = CreateCompatibleDC(0)
       AddPicture = mPicture
End Function

Private Sub MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim isSelect As Boolean
  Dim r As RECT
        
    If Not (m_NewObject Is Nothing) Or m_DrawingObject Then
         If DrawMode = MODE_EDITPOINT Then
            If Not Obj Is Nothing Then
              If Obj.EditPoint = False Then
                  'Redraw
                  frmVbDraw.drawToolbar.CheckButton 1, True
                  SelectTool 1
              Else
                 ChangeMenu
                 Exit Sub
              End If
            End If
         End If
         
    End If
    ' See where we clicked.
    
    If Not Obj Is Nothing Then isSelect = True
    
    If m_Rotate = False And m_Scale = False And m_Skew = False And m_Move = False Then
       Set Obj = FindObjectAt(X, Y)
    End If
    
    If (Obj Is Nothing) Then 'And m_SelectedObjects.Count <= 1 Then
        'Deselect all objects.
        DeselectAll
        m_Rotate = False
        m_Scale = False
        m_Move = False
        m_Skew = False
        LockObject = False
        ChangeMenu
        If isSelect = False Then Exit Sub
    Else
       If Button = 2 Then
           ViewMenu
          Exit Sub
       ElseIf Button = 1 Then
      
        ChangeMenu
        
        'See if the Shift key is pressed.
        If (Shift And vbShiftMask) Then
            ' Shift is pressed. Toggle this object's selection.
            If Obj.Selected Then
                DeselectVbdObject Obj
                m_Rotate = False
                m_Scale = False
                m_Move = False
                ClearBox
            Else
                SelectVbdObject Obj
                'GoTo MD1
            End If
        Else
            If m_SelectedObjects.Count > 1 Then 'And DrawMode <> MODE_EDITPOINT Then
               m_Move = True
               'm_EditObject = True
               m_StartX = X
               m_StartY = Y
               m_LastX = X
               m_LastY = Y
               Exit Sub
            End If
            ' Shift is not pressed. Select only this object.
            DeselectAllVbdObjects
md1:
            SelectVbdObject Obj
            
            LockObject = Obj.ObjLock
                                   
            DrawWidth = Obj.DrawWidth
            DrawStyle = Obj.DrawStyle
            FillStyle = Obj.FillStyle
            'Blend = Obj.Blend
            m_StartX = X
            m_StartY = Y
            m_LastX = X
            m_LastY = Y
            
            GetLimitBox
            
'            If (R.Left = 0 And R.Top = 0 And R.Right = 0 And R.Bottom = 0) Or Obj.TypeDraw = dPolydraw Then
'              Obj.Bound XminBox, YminBox, XmaxBox, YmaxBox
'            End If
            
            Xmin_Box = XminBox - m_StartX
            Xmax_Box = XmaxBox - m_StartX
            Ymin_Box = YminBox - m_StartY
            Ymax_Box = YmaxBox - m_StartY
            
            If m_Scale = False And m_Rotate = False And m_Skew = False Then 'And m_EditPoint = False Then
               'm_Move = True:
               'm_EditObject = True
            End If
            If m_Move Then
               PicCanvas.Line (m_StartX + Xmin_Box, m_StartY + Ymin_Box)-(m_StartX + Xmax_Box, m_StartY + Ymax_Box), , B
            End If
        End If
       End If
    End If
    
    If Not Obj Is Nothing Then
         RaiseEvent ColorSelected(1, Obj.FillColor)
         RaiseEvent ColorSelected(2, Obj.ForeColor)
    End If
    
    Redraw
        
    ' See if any objects are selected.
    RaiseEvent EnableMenusForSelection
End Sub

Private Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Ang As Single
Dim Msg As String

    If Not (Obj Is Nothing) Then
       If Obj.ObjLock = True Then GoTo NoSelect:
    End If
    
    'move point
     If Not (Obj Is Nothing) And Button = 1 And m_Move = True Then 'And m_EditObject Then
mm1:
      PicCanvas.DrawMode = vbInvert
      PicCanvas.DrawStyle = vbDot
          
      If m_Move Then
          PicCanvas.Line (m_LastX + Xmin_Box, m_LastY + Ymin_Box)-(m_LastX + Xmax_Box, m_LastY + Ymax_Box), vbRed, B
          m_LastX = X
          m_LastY = Y
          PicCanvas.Line (m_LastX + Xmin_Box, m_LastY + Ymin_Box)-(m_LastX + Xmax_Box, m_LastY + Ymax_Box), vbRed, B
          Msg = "Move X:" + Format(m_StartX, "0.0") + " Y:" + Format(m_StartY, "0.0") + " DX:" + Format(m_LastX, "0.0") + " DY:" + Format(m_LastY, "0.0")
          RaiseEvent MsgControl(Msg)
          Exit Sub
      End If
    
    'Scale point
    ElseIf Not (Obj Is Nothing) And Button = 1 And (m_Scale = True Or m_Skew = True) Then 'And m_EditObject Then
          PicCanvas.DrawMode = vbInvert
          PicCanvas.DrawStyle = vbDot
         
          Select Case m_ScaleType
          Case 1 'Left top Corner
                  PicCanvas.Line (m_LastX, m_LastY)-(XmaxBox, YmaxBox), , B
          Case 2 'Middle top
               If m_Scale Then
                  PicCanvas.Line (XminBox, m_LastY)-(XmaxBox, YmaxBox), , B
               Else
                  If XmaxBox - XminBox <= 0 Then Exit Sub
                  mDrawSkew (100 + ((m_LastX - m_StartX) * 100) / (XmaxBox - XminBox)), 100
               End If
          Case 3 'Right top Corner
                  PicCanvas.Line (XminBox, YmaxBox)-(m_LastX, m_LastY), , B
          Case 4 'Middle Right
               If m_Scale Then
                  PicCanvas.Line (XminBox, YminBox)-(m_LastX, YmaxBox), , B
               Else
                  If YmaxBox - YminBox <= 0 Then Exit Sub
                  mDrawSkew 100, (100 + ((m_StartY - m_LastY) * 100) / (YmaxBox - YminBox))
               End If
          Case 5 'Bottom Right corner
                  PicCanvas.Line (XminBox, YminBox)-(m_LastX + Xmax_Box, m_LastY + Ymax_Box), , B
          Case 6 'Middle Bottom
               If m_Scale Then
                  PicCanvas.Line (XminBox, YminBox)-(XmaxBox, m_LastY), , B
               Else
                  If XmaxBox - XminBox <= 0 Then Exit Sub
                  mDrawSkew (100 + ((m_StartX - m_LastX) * 100) / (XmaxBox - XminBox)), 100
               End If
          Case 7 'Left bottom corner
                  PicCanvas.Line (m_LastX, m_LastY)-(XmaxBox, YminBox), , B
          Case 8 'Middle left
               If m_Scale Then
                  PicCanvas.Line (m_LastX, YminBox)-(XmaxBox, YmaxBox), , B
               Else
                  If XmaxBox - XminBox <= 0 Then Exit Sub
                  mDrawSkew 100, (100 + ((m_LastY - m_StartY) * 100) / (YmaxBox - YminBox))
               End If
          End Select
          m_LastX = X
          m_LastY = Y
          Select Case m_ScaleType
            Case 1 'Left top Corner
                   PicCanvas.Line (m_LastX, m_LastY)-(XmaxBox, YmaxBox), , B
            Case 2 'Middle top
                If m_Scale Then
                   PicCanvas.Line (XminBox, m_LastY)-(XmaxBox, YmaxBox), , B
                Else
                   If XmaxBox - XminBox <= 0 Then Exit Sub
                   mDrawSkew (100 + ((m_LastX - m_StartX) * 100) / (XmaxBox - XminBox)), 100
                End If
            Case 3 'Right top Corner
                   PicCanvas.Line (XminBox, YmaxBox)-(m_LastX, m_LastY), , B
            Case 4 'Middle Right
                If m_Scale Then
                   PicCanvas.Line (XminBox, YminBox)-(m_LastX, YmaxBox), , B
                Else
                   If YmaxBox - YminBox <= 0 Then Exit Sub
                   mDrawSkew 100, (100 + ((m_StartY - m_LastY) * 100) / (YmaxBox - YminBox))
                End If
            Case 5 'Bottom Right corner
                   PicCanvas.Line (XminBox, YminBox)-(m_LastX + Xmax_Box, m_LastY + Ymax_Box), , B
            Case 6 'Middle Bottom
                If m_Scale Then
                   PicCanvas.Line (XminBox, YminBox)-(XmaxBox, m_LastY), , B
                Else
                   If XmaxBox - XminBox <= 0 Then Exit Sub
                   mDrawSkew (100 + ((m_StartX - m_LastX) * 100) / (XmaxBox - XminBox)), 100
                End If
            Case 7 'Left bottom corner
                   PicCanvas.Line (m_LastX, m_LastY)-(XmaxBox, YminBox), , B
            Case 8 'Middle left
                If m_Scale Then
                   PicCanvas.Line (m_LastX, YminBox)-(XmaxBox, YmaxBox), , B
                Else
                   If YmaxBox - YminBox <= 0 Then Exit Sub
                   mDrawSkew 100, (100 + ((m_LastY - m_StartY) * 100) / (YmaxBox - YminBox))
                End If
            End Select
            If m_Scale Then
               Msg = "Scale X:" + Format(m_StartX, "0.0") + " Y:" + Format(m_StartY, "0.0") + " DX:" + Format(m_LastX, "0.0") + " DY:" + Format(m_LastY, "0.0")
            ElseIf m_Skew Then
               Msg = "Skew X:" + Format(m_StartX, "0.0") + " Y:" + Format(m_StartY, "0.0") + " DX:" + Format(m_LastX, "0.0") + " DY:" + Format(m_LastY, "0.0")
            End If
            RaiseEvent MsgControl(Msg)
            Exit Sub

     'Rotate point
     ElseIf Not (Obj Is Nothing) And Button = 1 And m_Rotate = True Then 'And m_EditObject Then
            PicCanvas.DrawMode = vbInvert
            PicCanvas.DrawStyle = vbDot
            ' Create the rotation transformation.
            mDrawRotate m_LastX, m_LastY
            m_LastX = X
            m_LastY = Y
            mDrawRotate m_LastX, m_LastY
               
     'Change state and mousepointer
     ElseIf Not (Obj Is Nothing) Then ' And (m_DrawingObject Or m_EditObject) Then
        
'        If XminBox = 0 Then XminBox = X
'        If XmaxBox = 0 Then XmaxBox = X
'        If YminBox = 0 Then YminBox = Y
'        If YmaxBox = 0 Then YmaxBox = Y
         'Move object
        If XminBox + ((XmaxBox - XminBox) / 2) - GAP <= X And XminBox + ((XmaxBox - XminBox) / 2) + GAP >= X And _
           ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) >= Y And _
           Y >= ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) - GAP)) Then
            PicCanvas.MousePointer = 99
            PicCanvas.MouseIcon = ImageMouse(19).Picture
            GoSub State_Move
            Exit Sub
        End If
        
        'Point rotate
        If (XmaxBox + 18 <= X And XmaxBox + 22 >= X) And _
           ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) >= Y And _
           Y >= ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) - GAP)) Then
            PicCanvas.MousePointer = 99
            PicCanvas.MouseIcon = ImageMouse(12).Picture
            GoTo State_Rotate
            Exit Sub
        End If
        
        'If Obj.TypeDraw = dText Then GoTo NoSelect: Exit Sub
        
        'Left top Corner
        If (XminBox + GAP >= X And XminBox - GAP <= X) And _
           (YminBox + GAP >= Y And YminBox - GAP <= Y) Then
           m_ScaleType = 1
           m_Skew = False
           GoSub State_Scale
           If m_Scale Then
              PicCanvas.MousePointer = 8
            End If
            Exit Sub
        End If
        
        'Middle top
        If ((XminBox + (XmaxBox - XminBox) / 2 + GAP / 2) >= X And _
            ((XminBox + (XmaxBox - XminBox) / 2 + GAP / 2) - GAP) <= X) And _
           (YminBox >= Y And YminBox - GAP <= Y) Then
            m_ScaleType = 2
            m_Skew = True
             GoSub State_Scale
              PicCanvas.MousePointer = 99
              PicCanvas.MouseIcon = ImageMouse(13).Picture
            Exit Sub
        End If
                
       'Right top corner
        If (XmaxBox - GAP <= X And XmaxBox + GAP >= X) And _
           (YminBox - GAP <= Y And YminBox + GAP >= Y) Then
             m_ScaleType = 3
             m_Skew = False
            GoSub State_Scale
            If m_Scale Then
               PicCanvas.MousePointer = 6
            End If
            Exit Sub
        End If
                
        'Middle right
        If (XmaxBox - GAP <= X And XmaxBox + GAP >= X) And _
           ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) >= Y And _
           Y >= ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) - GAP)) Then
             m_ScaleType = 4
             m_Skew = True
             GoSub State_Scale
             PicCanvas.MousePointer = 99
             PicCanvas.MouseIcon = ImageMouse(14).Picture
            Exit Sub
        End If
        
        'Right botton corner
        If (XmaxBox <= X And X <= XmaxBox + GAP) And _
           (YmaxBox <= Y And Y <= YmaxBox + GAP) Then
             m_ScaleType = 5
             m_Skew = False
            GoSub State_Scale
            If m_Scale Then
               PicCanvas.MousePointer = 8
            End If
            Exit Sub
        End If
        
        'Middle botton
        If (XminBox + (XmaxBox - XminBox) / 2 + GAP / 2 >= X And _
            X >= (XminBox + (XmaxBox - XminBox) / 2 + GAP / 2) - GAP) And _
           (YmaxBox <= Y And YmaxBox + GAP >= Y) Then
            m_ScaleType = 6
            m_Skew = True
           GoSub State_Scale
             PicCanvas.MousePointer = 99
             PicCanvas.MouseIcon = ImageMouse(13).Picture
            Exit Sub
        End If
        
        'Botton left corner
        If (XminBox - GAP <= X And XminBox + GAP >= X) And _
           (YmaxBox - GAP <= Y And YmaxBox + GAP >= Y) Then
            m_ScaleType = 7
            m_Skew = False
            GoSub State_Scale
               PicCanvas.MousePointer = 6
            Exit Sub
        End If
                
        'Middle left
        If (XminBox - GAP <= X And XminBox + GAP >= X) And _
           ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) >= Y And _
           Y >= ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) - GAP)) Then
            m_Skew = True
            m_ScaleType = 8
            GoSub State_Scale
              PicCanvas.MousePointer = 99
              PicCanvas.MouseIcon = ImageMouse(14).Picture
            Exit Sub
        End If
        
       ' If m_ReadPenProperty = False And m_ReadFillProperty = False And Button = 0 Then
NoSelect:
          m_Scale = False
          m_Rotate = False
          m_Skew = False
          m_Move = False
          'PicCanvas.MousePointer = 99
          MouseIcon 1
          'PicCanvas.MouseIcon = ImageMouse(0).Picture
       ' End If
        
     'If not select state is move
     ElseIf Not (Obj Is Nothing) Then 'And m_EditObject Then
         
         m_Scale = False
         m_Rotate = False
         m_Skew = False
         m_Move = True
     End If
    
    Exit Sub
    
State_Scale:
   If m_Skew = False And m_Scale = False Then m_Scale = True
   If m_Skew = True Then
      m_Scale = False
      m_Skew = True
    Else
      m_Scale = True
      m_Skew = False
    End If
    m_Move = False
    m_Rotate = False
    'm_EditObject = True
    Return
Exit Sub

State_Rotate:
    m_Scale = False
    m_Skew = False
    m_Move = False
    m_Rotate = True
    'm_EditObject = True
Exit Sub

State_Move:
    m_Scale = False
    m_Skew = False
    m_Move = True
    m_Rotate = False
    'm_EditObject = True
    Return

End Sub

Private Sub MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
Dim X1 As Single
Dim X2 As Single
Dim Y1 As Single
Dim Y2 As Single
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim Ang As Single
Dim Msg As String
Dim r As RECT
          
     If m_SelectedObjects.Count > 1 And m_Move Then
          m_LastX = X
          m_LastY = Y
     End If
     
     
     If Not (Obj Is Nothing) Then
     
'MU1:
      If Obj.ObjLock = True Then GoTo NoSelect
       PicCanvas.DrawMode = vbCopyPen
       
       If m_Move And Button = 1 Then
          'If (X - m_StartX) / m_ZoomFactor = 0 And (Y - m_StartY) / m_ZoomFactor = 0 Then Exit Sub
          If (X - m_StartX) = 0 And (Y - m_StartY) = 0 Then Exit Sub
             PicCanvas.Line (m_LastX + Xmin_Box, m_LastY + Ymin_Box)-(m_LastX + Xmax_Box, m_LastY + Ymax_Box), , B
             'TransformPoint (X - m_StartX) / m_ZoomFactor, (Y - m_StartY) / m_ZoomFactor
             'Msg = "Move X:" + Format((X - m_StartX) / m_ZoomFactor, "0.0") + _
                       " Y:" + Format((Y - m_StartY) / m_ZoomFactor, "0.0")
             TransformPoint (X - m_StartX), (Y - m_StartY)
             Msg = "Move X:" + Format((X - m_StartX), "0.0") + _
                       " Y:" + Format((Y - m_StartY), "0.0")
             RaiseEvent MsgControl(Msg)
 
       ElseIf m_Rotate And Button = 1 Then
            m_LastX = X '/ m_ZoomFactor
            m_LastY = Y '/ m_ZoomFactor
            Ang = m2GetAngle3P(XminBox + (XmaxBox - XminBox) / 2, YminBox + (YmaxBox - YminBox) / 2, _
                               XminBox + (XmaxBox - XminBox), m_StartY, _
                               m_LastX, m_LastY)
            'TransformRotate Ang, XminBox / m_ZoomFactor, YminBox / m_ZoomFactor, XmaxBox / m_ZoomFactor, YmaxBox / m_ZoomFactor
            TransformRotate Ang, XminBox, YminBox, XmaxBox, YmaxBox
            Msg = "Rotate angle:" + Format(Ang, "0.0")
            RaiseEvent MsgControl(Msg)
             
       ElseIf m_Scale Or m_Skew And Button = 1 Then
           Select Case m_ScaleType
           Case 1 'Left top Corner
              If (XmaxBox - XminBox) <> 0 And (YmaxBox - YminBox) <> 0 Then
                 If m_Scale Then
                     TransformScale (100 + ((m_StartX - X) * 100) / (XmaxBox - XminBox)), _
                                    (100 + ((m_StartY - Y) * 100) / (YmaxBox - YminBox))
                 End If
              End If
           Case 2 'Middle top
              If (YmaxBox - YminBox) <> 0 Then
                 If m_Scale Then
                    TransformScale 100, _
                                   (100 + ((m_StartY - Y) * 100) / (YmaxBox - YminBox))
                 Else
                    TransformSkew (100 + ((X - m_StartX) * 100) / (XmaxBox - XminBox)), _
                                  100
                 End If
              End If
           Case 3 'Right top Corner
              If (XmaxBox - XminBox) <> 0 And (YmaxBox - YminBox) <> 0 Then
                 If m_Scale Then
                    TransformScale (100 + ((X - m_StartX) * 100) / (XmaxBox - XminBox)), _
                                   (100 + ((m_StartY - Y) * 100) / (YmaxBox - YminBox))
                 End If
              End If
           Case 4 'Middle Right
             If (XmaxBox - XminBox) <> 0 Then
                If m_Scale Then
                   TransformScale (100 + ((X - m_StartX) * 100) / (XmaxBox - XminBox)), _
                                  100
                Else
                   TransformSkew 100, _
                                 (100 + ((m_StartY - Y) * 100) / (YmaxBox - YminBox))
                End If
             End If
           Case 5 'Bottom Right corner
             If (XmaxBox - XminBox) <> 0 And (YmaxBox - YminBox) <> 0 Then
                If m_Scale Then
                   TransformScale (100 + ((X - m_StartX) * 100) / (XmaxBox - XminBox)), _
                                  (100 + ((Y - m_StartY) * 100) / (YmaxBox - YminBox))
                End If
             End If
           Case 6 'Middle Bottom
             If (YmaxBox - YminBox) <> 0 Then
                If m_Scale Then
                   TransformScale 100, _
                                  (100 + ((Y - m_StartY) * 100) / (YmaxBox - YminBox))
                Else
                   TransformSkew (100 + ((m_StartX - X) * 100) / (XmaxBox - XminBox)), _
                                   100
                End If
             End If
           Case 7 'Left bottom corner
             If (XmaxBox - XminBox) <> 0 And (YmaxBox - YminBox) <> 0 Then
                If m_Scale Then
                   TransformScale (100 + ((m_StartX - X) * 100) / (XmaxBox - XminBox)), _
                                  (100 + ((Y - m_StartY) * 100) / (YmaxBox - YminBox))
                End If
             End If
           Case 8 'Middle left
             If (XmaxBox - XminBox) <> 0 Then
                If m_Scale Then
                   TransformScale (100 + ((m_StartX - X) * 100) / (XmaxBox - XminBox)), _
                                   100
                Else
                   TransformSkew 100, _
                                 (100 + ((Y - m_StartY) * 100) / (YmaxBox - YminBox))
                End If
             End If
           End Select
       End If
       Redraw
       GoTo NoSelect
     End If
    
     SelectTool 1 '"Arrow"
NoSelect:
     m_Move = False
     m_Rotate = False
     m_Scale = False
     m_Skew = False
     If Not Obj Is Nothing Then
         GetRgnBox Obj.hRegion, r
         XminBox = r.Left
         YminBox = r.Top
         XmaxBox = r.Right
         YmaxBox = r.Bottom
'         If Obj.TypeDraw = dPolydraw Then
'            Obj.Bound XminBox, YminBox, XmaxBox, YmaxBox
'         End If
     End If
  
End Sub

'
Private Sub MouseDownPoint(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ObjPoint As vbdObject
    Dim mStep As Single, TP As Integer
    Dim Points() As PointAPI
    Dim isSelect As Boolean
    
    mStep = IIf(GAP / gZoomFactor < 0.5, 1, GAP / gZoomFactor) ', GAP)
    
    If PicCanvas.ToolTipText = "" Then 'no select node
      Set ObjPoint = FindObjectAt(X, Y)
      If Not ObjPoint Is Nothing And Not Obj Is Nothing Then
        If Obj Is ObjPoint = False Then
SelObj1:
           DeselectAllVbdObjects
           Set Obj = ObjPoint
           SelectVbPoint Obj
           Obj.ReadTrPoint m_NumPoints, MX(), mY(), m_TypePoint
           Obj.NewPoint m_NumPoints, MX, mY, m_TypePoint
           ReDim m_OriginalPoints(1 To m_NumPoints)
           For r = 1 To m_NumPoints
              m_OriginalPoints(r).X = MX(r)
              m_OriginalPoints(r).Y = mY(r)
           Next
           Redraw
           Set ObjPoint = Nothing
        Else
           DeselectAllVbdObjects
           If Not Obj Is Nothing Then
           SelectVbPoint Obj
           End If
           Redraw
        End If
        
      ElseIf Not ObjPoint Is Nothing And Obj Is Nothing Then
         GoTo SelObj1
         'Stop
      ElseIf (ObjPoint Is Nothing And Obj Is Nothing) Or (ObjPoint Is Nothing And Not Obj Is Nothing) Then
         DeselectAll
         ChangeMenu
         If Not ObjPoint Is Nothing Or Not Obj Is Nothing Then isSelect = False Else isSelect = True
         Set Obj = Nothing
         PicCanvas.ToolTipText = ""
         If isSelect Then
            Redraw
         End If
         
         Exit Sub
         
      End If
    ElseIf Not Obj Is Nothing Then  'select point
      If Obj.EditPoint = False Then Obj.EditPoint = True
    Else
       Exit Sub
    End If
    
    If Obj Is Nothing Then ' no select object
       Set Obj = FindObjectAt(X, Y)
       If Not Obj Is Nothing Then 'select object
          If Obj.ObjLock = True Then Set Obj = Nothing: Exit Sub
         
          SelectVbPoint Obj
          
          Obj.ReadTrPoint m_NumPoints, MX(), mY(), m_TypePoint
          Obj.NewPoint m_NumPoints, MX, mY, m_TypePoint
          ReDim m_OriginalPoints(1 To m_NumPoints)
          For r = 1 To m_NumPoints
              m_OriginalPoints(r).X = MX(r)
              m_OriginalPoints(r).Y = mY(r)
          Next
          Redraw
          DrawPoint
       Else 'no select object
          Redraw
          Exit Sub
       End If
    End If
    
    GetLimitBox
    If Obj.TypeDraw = dPolygon Then
         e_NumPoints = 2
    ElseIf Obj.TypeDraw = dRectAngle Then
         e_NumPoints = 4
    Else
         e_NumPoints = m_NumPoints
    End If
    
    m_SelectPoint = 0
    PicCanvas.ToolTipText = ""
    'find point
    For i = 1 To e_NumPoints
        If X >= m_OriginalPoints(i).X - mStep And X <= m_OriginalPoints(i).X + mStep And _
           Y >= m_OriginalPoints(i).Y - mStep And Y <= m_OriginalPoints(i).Y + mStep Then
           m_SelectPoint = i
           PicCanvas.ToolTipText = "Select Point:" + Str(m_SelectPoint)
           Exit For
        End If
    Next
    
    'menu point
    If Button = 2 And m_SelectPoint > 0 And (Obj.TypeDraw = dPolydraw Or _
                                             Obj.TypeDraw = dScribble Or _
                                             Obj.TypeDraw = dCalligraphic) Then
        Redraw
        SelectMenu = MenuNode
          If SelectMenu = 1 Then  'Add Node
                mAddNode X, Y
            ElseIf SelectMenu = 2 Then  'Delete Node
                mDeleteNode
            ElseIf SelectMenu = 4 Then 'Line
                mtoLine
            ElseIf SelectMenu = 5 Then 'Curve
                  mtoCurve
            ElseIf SelectMenu = 7 Then 'Auto Close
                mCloseNode
            ElseIf SelectMenu = 8 Then 'Auto open
                mOpenNode
            ElseIf SelectMenu = 10 Then 'Break
                mBreakNode
            End If
            
            Obj.NewPoint m_NumPoints, MX, mY, m_TypePoint
            Obj.ReadTrPoint m_NumPoints, MX, mY, m_TypePoint
            If m_SelectPoint > m_NumPoints Then m_SelectPoint = m_NumPoints
            Redraw
            Exit Sub
     End If
        
      Redraw
      
      m_LastX = X
      m_LastY = Y
      
      If m_SelectPoint > 0 Then
          m_DrawingPoint = True
          PicCanvas.DrawMode = vbInvert
          PicCanvas.DrawStyle = vbDot
          GetLimitBox
          If Obj.TypeDraw = dPolygon Then
              m_NumPoints = m_NumPoints - 1
                 Points = CreatePolygon2(m_OriginalPoints(m_SelectPoint).X, m_OriginalPoints(m_SelectPoint).Y, _
                                         m_OriginalPoints(m_SelectPoint + 2).X, m_OriginalPoints(m_SelectPoint + 2).Y, _
                                         m_NumPoints / 2, m_LastX, m_LastY)
                iCounter = 0
                For i = m_SelectPoint To m_NumPoints Step 2
                    iCounter = iCounter + 1
                    m_OriginalPoints(i) = Points(iCounter)
                Next
            m_NumPoints = m_NumPoints + 1
            m_OriginalPoints(m_NumPoints) = m_OriginalPoints(1)
          Else
             Call PolyPoints(m_SelectPoint, m_LastX, m_LastY)
'             Stop
             
'             If m_TypePoint(m_NumPoints) = 3 And Obj.TypeDraw = dRectAngle Then
'                m_OriginalPoints(m_NumPoints).X = m_OriginalPoints(1).X
'                m_OriginalPoints(m_NumPoints).Y = m_OriginalPoints(1).Y
'             End If
          End If
      End If
      
     DrawPoint
End Sub
Private Sub MouseMovePoint(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Dim Points() As PointAPI, iCounter As Integer
        
        If m_DrawingPoint = False Then Exit Sub
      
        DrawPoint
        
        ' Update the point.
        If Obj.TypeDraw = dPolygon Then
             
             m_NumPoints = m_NumPoints - 1
             Points = CreatePolygon(m_LastX, m_LastY, X, Y, m_NumPoints / 2)
             iCounter = 0
             For i = m_SelectPoint To m_NumPoints Step 2
                  iCounter = iCounter + 1
                  m_OriginalPoints(i) = Points(iCounter)
              Next
              m_NumPoints = m_NumPoints + 1
             m_OriginalPoints(m_NumPoints) = m_OriginalPoints(1)
             
         ElseIf Obj.TypeDraw = dRectAngle Then
            m_LastX = X
            m_LastY = Y
            Call PolyPoints(m_SelectPoint, m_LastX, m_LastY)
            Select Case m_SelectPoint
            Case 1
               m_OriginalPoints(4).X = m_LastX
               m_OriginalPoints(2).Y = m_LastY
            Case 2
               m_OriginalPoints(3).X = m_LastX
               m_OriginalPoints(1).Y = m_LastY
            Case 3
               m_OriginalPoints(2).X = m_LastX
               m_OriginalPoints(4).Y = m_LastY
            Case 4
               m_OriginalPoints(1).X = m_LastX
               m_OriginalPoints(3).Y = m_LastY
            End Select
            
         Else
           m_LastX = X
           m_LastY = Y
           Call PolyPoints(m_SelectPoint, m_LastX, m_LastY)
         End If
         DrawPoint
       
End Sub

Private Sub MouseUpPoint(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Dim Points() As PointAPI, iCounter As Integer
       
        If m_DrawingPoint = False Or Obj Is Nothing Then Exit Sub
          
         PicCanvas.DrawMode = vbCopyPen
         PicCanvas.DrawStyle = vbSolid

         If Obj.TypeDraw = dPolygon Then
             
             m_NumPoints = m_NumPoints - 1
             Points = CreatePolygon(m_LastX, m_LastY, X, Y, m_NumPoints / 2)
             iCounter = 0
             For i = m_SelectPoint To m_NumPoints Step 2
                 iCounter = iCounter + 1
                  m_OriginalPoints(i) = Points(iCounter)
              Next
              m_NumPoints = m_NumPoints + 1
              m_OriginalPoints(m_NumPoints) = m_OriginalPoints(1)
         Else
             m_LastX = X
             m_LastY = Y
             Call PolyPoints(m_SelectPoint, m_LastX, m_LastY)
         End If
         
         DrawPoint
         m_DrawingPoint = False
          
          ReDim MX(1 To m_NumPoints)
          ReDim mY(1 To m_NumPoints)
          For r = 1 To m_NumPoints
              MX(r) = m_OriginalPoints(r).X
              mY(r) = m_OriginalPoints(r).Y
          Next
         Obj.NewPoint m_NumPoints, MX, mY, m_TypePoint
         Obj.ReadTrPoint m_NumPoints, MX, mY, m_TypePoint
         
         Redraw
         DrawPoint
         Set_Dirty
        ' DrawPoint
        PicCanvas.ToolTipText = ""
        RaiseEvent MsgControl(PicCanvas.ToolTipText)
        m_SelectPoint = 0
   
End Sub

Private Sub mCreatePolygon(ByVal X As Single, ByVal Y As Single)
        Dim X1 As Single, Y1 As Single
       ' GetLimitBox
        X1 = Sqr((XminBox + (XmaxBox - XminBox) / 2 - X) * (XminBox + (XmaxBox - XminBox) / 2 - X))
        Y1 = Sqr((YminBox + (YmaxBox - YminBox) / 2 - Y) * (YminBox + (YmaxBox - YminBox) / 2 - Y))
'        'Debug.Print "L:" + Str(X) + "," + Str(Y) + "-" + Str(XminBox + (XmaxBox - XminBox) / 2) + "," + Str(YminBox + (YmaxBox - YminBox) / 2) + " - c:" + Str(X + X1) + "," + Str(Y + Y1)
        m_NumPoints = m_NumPoints - 1
        m_OriginalPoints = CreatePolygon(XminBox + (XmaxBox - XminBox) / 2, YminBox + (YmaxBox - YminBox) / 2, X, Y, m_NumPoints)
        m_NumPoints = m_NumPoints + 1
        ReDim Preserve m_OriginalPoints(1 To m_NumPoints)
        m_OriginalPoints(m_NumPoints) = m_OriginalPoints(1)
End Sub

'Popup Menu for point
Private Function MenuNode() As Long
    Dim PT As PointAPI
    Dim ret As Long
    Dim wFlag0 As Long, wFlag1 As Long, wFlag2 As Long, wFlag3 As Long, wFlag4 As Long, wFlag5 As Long, wFlag6 As Long
    
    If IsControl(m_TypePoint, m_SelectPoint) = True Then
       wFlag0 = MF_GRAYED Or MF_DISABLED
    Else
        wFlag0 = MF_STRING
    End If
    If m_SelectPoint = 0 Or m_SelectPoint = 1 Or IsControl(m_TypePoint, m_SelectPoint) = True Then
        wFlag1 = MF_GRAYED Or MF_DISABLED
    Else
        wFlag1 = MF_STRING
    End If
    
    If m_SelectPoint > 0 Then
       If m_TypePoint(m_SelectPoint) = 6 Then
          If m_SelectPoint + 1 > m_NumPoints Then
               wFlag2 = MF_GRAYED Or MF_DISABLED
              wFlag3 = MF_GRAYED Or MF_DISABLED
          Else
          If m_TypePoint(m_SelectPoint + 1) = 2 Then wFlag2 = MF_GRAYED Or MF_DISABLED Else wFlag2 = MF_STRING
          If m_TypePoint(m_SelectPoint + 1) = 4 Then wFlag3 = MF_GRAYED Or MF_DISABLED Else wFlag3 = MF_STRING
          End If
       ElseIf m_TypePoint(m_SelectPoint) = 3 Then
           wFlag2 = MF_GRAYED Or MF_DISABLED
           wFlag3 = MF_GRAYED Or MF_DISABLED
       Else
          If m_SelectPoint >= m_NumPoints Then m_SelectPoint = m_SelectPoint - 1
          If m_TypePoint(m_SelectPoint + 1) = 2 Or _
             m_TypePoint(m_SelectPoint + 1) = 6 Or _
             m_TypePoint(m_SelectPoint + 1) = 3 Or _
             IsControl(m_TypePoint, m_SelectPoint) = True Then
              wFlag2 = MF_GRAYED Or MF_DISABLED
          Else
              wFlag2 = MF_STRING
          End If
          
          If m_TypePoint(m_SelectPoint + 1) = 4 Or _
             m_TypePoint(m_SelectPoint + 1) = 6 Or _
             m_TypePoint(m_SelectPoint + 1) = 3 Or _
             IsControl(m_TypePoint, m_SelectPoint) = True Then
             wFlag3 = MF_GRAYED Or MF_DISABLED
          Else
             wFlag3 = MF_STRING
          End If
                    
       End If
    Else
       wFlag2 = MF_GRAYED Or MF_DISABLED
       wFlag3 = MF_GRAYED Or MF_DISABLED
    End If
    
    If m_TypePoint(m_NumPoints) = 3 Or m_TypePoint(m_NumPoints) = 5 Then wFlag4 = MF_GRAYED Or MF_DISABLED Else wFlag4 = MF_STRING
    
    If IsOpening(m_TypePoint) Then wFlag5 = MF_STRING Else wFlag5 = MF_GRAYED Or MF_DISABLED
    
    If m_SelectPoint > 0 And m_SelectPoint <> m_NumPoints And m_TypePoint(m_NumPoints) <> 6 And _
       IsControl(m_TypePoint, m_SelectPoint) = False Then
       wFlag6 = MF_STRING
    Else
       wFlag6 = MF_GRAYED Or MF_DISABLED
    End If
    
    hMenu = CreatePopupMenu()
    AppendMenu hMenu, wFlag0, 1, "Add node(s)"
    AppendMenu hMenu, wFlag1, 2, "Delete node(s)" + vbTab + "(Del)"
    AppendMenu hMenu, MF_SEPARATOR, 3, ByVal 0&
    AppendMenu hMenu, wFlag2, 4, "To Line"
    AppendMenu hMenu, wFlag3, 5, "To Curve" '
    AppendMenu hMenu, MF_SEPARATOR, 6, ByVal 0&
    AppendMenu hMenu, wFlag4, 7, "Auto Close"
    AppendMenu hMenu, wFlag5, 8, "Auto Open"
    AppendMenu hMenu, MF_SEPARATOR, 9, ByVal 0&
    AppendMenu hMenu, wFlag6, 10, "Break node"
    AppendMenu hMenu, MF_GRAYED Or MF_DISABLED, 11, "Break Apart"
    
    GetCursorPos PT
    ret = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, PT.X, PT.Y, PicCanvas.hWnd, ByVal 0&)
    DestroyMenu hMenu
    MenuNode = ret
End Function

'Read Color property
Private Sub ReadColorProperty(ByVal X As Single, ByVal Y As Single)
       Dim rOBJ As vbdObject
       
       Set rOBJ = FindObjectAt(X, Y)
       If Not rOBJ Is Nothing Then
          Select Case DrawingMode
          Case MODE_ReadFill
               CtlFill1.Color1 = rOBJ.FillColor
               CtlFill1.Color2 = rOBJ.FillColor2
               CtlFill1.FillStyle = rOBJ.FillStyle '+ 1
               CtlFill1.NamePattern = rOBJ.Pattern
               CtlFill1.TypeGradient = rOBJ.Gradient
               CtlFill1.Blend = rOBJ.Blend
          Case MODE_ReadPen
               PicPenColor.BackColor = rOBJ.ForeColor
               ' Select the 1 pixel DrawWidth.
               icbDrawWidth.SelectedItem = icbDrawWidth.ComboItems(rOBJ.DrawWidth)
               icbDrawWidth.ToolTipText = icbDrawWidth.ComboItems(rOBJ.DrawWidth).Key
               ' Select the solid DrawStyle.
               icbDrawStyle.SelectedItem = icbDrawStyle.ComboItems(rOBJ.DrawStyle + 1)
               icbDrawStyle.ToolTipText = icbDrawStyle.ComboItems(rOBJ.DrawStyle + 1).Key
          End Select
       End If
       
       Set rOBJ = Nothing
       DrawingMode = OldDrawingMode
       MouseIcon DrawingMode
End Sub

Private Sub cmdSysColorsPen_Click()
     OpenColorDialog PicPenColor
End Sub

Private Sub ComCorner_Click()
    If (HScroll1.Min - HScroll1.Max) \ 2 >= HScroll1.Min And (HScroll1.Min - HScroll1.Max) \ 2 <= HScroll1.Max Then
        HScroll1.Value = (HScroll1.Min - HScroll1.Max) \ 2
    End If
    If (VScroll1.Max - VScroll1.Min) \ 2 >= VScroll1.Min And (VScroll1.Max - VScroll1.Min) \ 2 < VScroll1.Max Then
        VScroll1.Value = (VScroll1.Max - VScroll1.Min) \ 2
    End If
End Sub

Private Sub ComDropperPen_Click()
      'm_ReadPenProperty = True
      OldDrawingMode = DrawingMode
      DrawingMode = MODE_ReadPen
      SelectTool 20 '"DropperFill"
End Sub

Private Sub Command1_Click()
     Dim id As Long
     If Val(LabelId) <= UBound(m_OriginalPoints) Then
        id = Val(LabelId)
        m_OriginalPoints(id).X = Text1.Text
        m_OriginalPoints(id).Y = Text2.Text
     End If
End Sub

Private Sub CommandPen_Click()
     DrawStyle = icbDrawStyle.SelectedItem.Index - 1
     DrawWidth = icbDrawWidth.SelectedItem.Index
     ForeColor = PicPenColor.BackColor
     Redraw
End Sub

Private Sub CtlFill1_Apply(nTypeFill As Integer, nFillStyle As Integer, nColor1 As Long, nColor2 As Long, nPattern As String, nTypeGradient As Integer, mBlend As Integer)
      Select Case nTypeFill
      Case 1
            FillStyle = nFillStyle  'icbFillStyle.SelectedItem.Index - 1
            FillColor = nColor1
      Case 2
            FillStyle = nFillStyle
            FillColor = nColor1
            FillColor2 = nColor2
            Pattern = nPattern
      Case 3
            FillStyle = nFillStyle
            FillColor = nColor1
            FillColor2 = nColor2
            Gradient = nTypeGradient
      Case 4
            FillStyle = nFillStyle
            Pattern = nPattern
      End Select
      
      Blend = mBlend
      Redraw
End Sub


Private Sub CtlFill1_ApplyImage(nTypeFill As Integer, nFillStyle As Integer, nPattern As String, nPicture As stdole.StdPicture, mBlend As Integer)
       Dim nobj As vbdObject
       
       If Obj Is Nothing Then Exit Sub
        FillStyle = nFillStyle
        Pattern = nPattern
        Blend = mBlend
        Set nobj = Obj
        
        ' Add the transformation to the selected objects.
        For Each nobj In m_SelectedObjects
           If nobj.Selected = True Then
              nobj.Pattern = nPattern
              Set nobj.Picture = nPicture
           End If
        Next nobj
        Redraw
        Set nobj = Nothing
End Sub

Private Sub CtrTranform1_TransformMove(X_Move As Single, Y_Move As Single)
       TransformPoint X_Move, Y_Move
End Sub

Private Sub CtrTranform1_TransformMirror(X_Skew As Integer, Y_Skew As Integer)
       TransformMirror X_Skew, Y_Skew
End Sub

Private Sub CtrTranform1_TransformRotate(t_Angle As Single, xmin As Single, ymin As Single, xmax As Single, ymax As Single)
       If Obj Is Nothing Then Exit Sub
       Obj.Bound XminBox, YminBox, XmaxBox, YmaxBox
       TransformRotate t_Angle, XminBox, YminBox, XmaxBox, YmaxBox
End Sub

Private Sub CtrTranform1_TransformScale(X_scale As Single, Y_Scale As Single)
       m_ScaleType = 9
       TransformScale X_scale, Y_Scale
       m_ScaleType = 0
End Sub

Private Sub CtrTranform1_TransformSkew(X_Skew As Single, Y_Skew As Single)
       m_ScaleType = 9
       TransformSkew X_Skew, Y_Skew
       m_ScaleType = 0
End Sub

Private Sub HScroll1_Change()
' On Error Resume Next
'    'picCanvas.Left = HScroll1.Value
'    m_CanvasLeft = HScroll1.Value
'    PicCanvas.Visible = False
'    PicCanvas.Left = m_CanvasLeft
'    ReDrawPage
'    If mLockControl = False Then PicCanvas.Visible = True
     If IgnoreSbarChange Then Exit Sub
   ' 'Debug.Print "HScrollBar.Value", HScroll1.Value
    HScrollBarChanged
End Sub

Private Sub ReDrawPage()
    Dim SCR As Single
     If mLockControl = True Then Exit Sub
   
    SCR = Round((DataYmax - DataYmin) / (Wymax - Wymin), 3)
    SCR = 20 / SCR
'    If ScR >= 60 Then ScR = ScR / 4
   PicCanvas.BackColor = RGB(240, 240, 240)
    PicCanvas.Line (SCR, SCR)-(m_CanvasWidth + SCR, m_CanvasHeight + SCR), QBColor(7), BF
    PicCanvas.Line (0, 0)-(m_CanvasWidth, m_CanvasHeight), m_BackColor, BF
    PicCanvas.Line (0, 0)-(m_CanvasWidth, m_CanvasHeight), , B
    
End Sub


Private Sub ReDrawRuler(mOrientation As Integer)
    
    Dim mySmallScale As Single
    Dim myValue As String, RSm As String
    Dim i As Single, mvarScale As Long, j As Single, Div As Integer
    Dim SCR As Single, Oldsize As Single, OldFORECOLOR As Long, TW As Single
    
    Oldsize = PicCanvas.Font.Size
    OldFORECOLOR = PicCanvas.ForeColor
    PicCanvas.ForeColor = RGB(0, 0, 0)
    PicCanvas.Font.Name = "Small fonts"
    PicCanvas.Font.Size = 6
    PicCanvas.Font.Italic = False
    PicCanvas.Font.Bold = False
    PicCanvas.Font.Strikethrough = False
    PicCanvas.Font.Underline = False
    
     'Set scaling
    Select Case gScaleMode
        Case 1 'smTwips
            mvarScale = 1000
            RSm = "Tw"
            Div = 10
        Case 3 'smPixels
            mvarScale = (Screen.TwipsPerPixelX * 100) / 1.5
            Div = 100
            RSm = "Pix"
        Case 6 'smMilimeters
            mvarScale = 5670 / 1.5
            Div = 100
            RSm = "mm"
        Case 5 'smInches
            mvarScale = 1440 / 1.5
            Div = 1
            RSm = "In"
        Case Else
           Exit Sub
    End Select
    
    mvarScale = mvarScale / 10
    mySmallScale = mvarScale / 10
    
    'small zoom
    If gZoomFactor < 1 Then
       mvarScale = mvarScale * 10
       mySmallScale = mySmallScale * 10
       Div = Div * 10
    ElseIf gZoomFactor > 10 Then
'       mvarScale = mvarScale / 2
'       mySmallScale = mySmallScale / 2
    End If
    TW = PicCanvas.TextWidth("Pix")
    SCR = Round((DataYmax - DataYmin) / (Wymax - Wymin), 3)
    SCR = 50 / SCR
    If TW > SCR Then
       SCR = TW * 2
       'Div = Div / 100
    ElseIf TW * 2 < SCR Then
       SCR = TW * 2
      ' Div = Div / 100
    End If
    Select Case mOrientation
    Case 1 'Horizontal
        CurrentX = 0
        PicCanvas.Line (PicCanvas.ScaleLeft, PicCanvas.ScaleTop)-(PicCanvas.ScaleLeft + PicCanvas.ScaleWidth, PicCanvas.ScaleTop + SCR), vbButtonFace, BF
        PicCanvas.Line (PicCanvas.ScaleLeft + 1 / gZoomFactor, PicCanvas.ScaleTop + 1 / gZoomFactor)-(PicCanvas.ScaleLeft + PicCanvas.ScaleWidth - 1 / gZoomFactor, PicCanvas.ScaleTop + SCR - 1 / gZoomFactor), RGB(150, 150, 150), B
        PicCanvas.Line (PicCanvas.ScaleLeft + 1 / gZoomFactor, PicCanvas.ScaleTop + 1 / gZoomFactor)-(PicCanvas.ScaleLeft + PicCanvas.ScaleWidth, PicCanvas.ScaleTop + SCR), RGB(255, 255, 255), B
        PicCanvas.Line (PicCanvas.ScaleLeft, PicCanvas.ScaleTop)-(PicCanvas.ScaleLeft + PicCanvas.ScaleWidth, PicCanvas.ScaleTop + SCR), , B
        If PicCanvas.ScaleLeft > 0 Then StartLeft = PicCanvas.ScaleLeft Else StartLeft = 0
        For j = StartLeft To PicCanvas.ScaleLeft + PicCanvas.ScaleWidth Step mvarScale
            'Draw big line
            PicCanvas.Line (j, PicCanvas.ScaleTop)-(j, PicCanvas.ScaleTop + SCR)
             'Print Value
            myValue = Round((j / mvarScale) * Div, 1)
            PicCanvas.CurrentY = PicCanvas.ScaleTop
            If j = 0 Then myValue = myValue + RSm
            PicCanvas.Print myValue
            'Draw small lines
            n = 0
            For i = j + mySmallScale To j + mvarScale Step mySmallScale
                n = n + 1
                If n = 5 Then
                    PicCanvas.Line (i, PicCanvas.ScaleTop + (SCR / 2))-(i, PicCanvas.ScaleTop + SCR)
                Else
                    PicCanvas.Line (i, PicCanvas.ScaleTop + (SCR * 0.6))-(i, PicCanvas.ScaleTop + SCR)
                End If
                If gScaleMode = 6 Then
                   For t = i - mySmallScale To i + mySmallScale Step (mySmallScale / 10)
                      PicCanvas.Line (t, PicCanvas.ScaleTop + (SCR * 0.6))-(t, PicCanvas.ScaleTop + SCR)
                   Next
                   PicCanvas.Line (i, PicCanvas.ScaleTop)-(i, PicCanvas.ScaleTop + SCR)
                   myValue = Round((i / mvarScale) * Div)
                   PicCanvas.CurrentY = PicCanvas.ScaleTop
                   PicCanvas.Print myValue
                End If
            Next i
        Next j
        
        For j = 0 To PicCanvas.ScaleLeft - 1000 Step -mvarScale
            'Draw big line
            PicCanvas.Line (j, PicCanvas.ScaleTop)-(j, PicCanvas.ScaleTop + SCR)
            'Print Value
            myValue = Round((j / mvarScale) * Div)
            CurrentY = 0
            PicCanvas.CurrentY = PicCanvas.ScaleTop
            PicCanvas.Print myValue
            'Draw small lines
            n = 0
            For i = j + -mySmallScale To j + -mvarScale Step -mySmallScale
                n = n + 1
                If n = 5 Then
                    PicCanvas.Line (i, PicCanvas.ScaleTop + (SCR / 2))-(i, PicCanvas.ScaleTop + SCR)
                Else
                    PicCanvas.Line (i, PicCanvas.ScaleTop + (SCR * 0.6))-(i, PicCanvas.ScaleTop + SCR)
                End If
                If gScaleMode = 6 Then
                   For t = j To i Step -(mySmallScale / 10)
                      PicCanvas.Line (t, PicCanvas.ScaleTop + (SCR * 0.6))-(t, PicCanvas.ScaleTop + SCR)
                   Next
                   PicCanvas.Line (i, PicCanvas.ScaleTop)-(i, PicCanvas.ScaleTop + SCR)
                   myValue = Round((i / mvarScale) * Div)
                   PicCanvas.CurrentY = PicCanvas.ScaleTop
                   PicCanvas.Print myValue
                End If
            Next i
        Next j

    Case 2 'Vertical
        PicCanvas.Line (PicCanvas.ScaleLeft, PicCanvas.ScaleTop)-(PicCanvas.ScaleLeft + SCR, PicCanvas.ScaleTop + PicCanvas.ScaleHeight), vbButtonFace, BF
        PicCanvas.Line (PicCanvas.ScaleLeft + 1 / gZoomFactor, PicCanvas.ScaleTop + 1 / gZoomFactor)-(PicCanvas.ScaleLeft + SCR - 1 / gZoomFactor, PicCanvas.ScaleTop + PicCanvas.ScaleHeight - 1 / gZoomFactor), RGB(150, 150, 150), B
        PicCanvas.Line (PicCanvas.ScaleLeft + 1 / gZoomFactor, PicCanvas.ScaleTop + 1 / gZoomFactor)-(PicCanvas.ScaleLeft + SCR, PicCanvas.ScaleTop + PicCanvas.ScaleHeight), RGB(255, 255, 255), B
        PicCanvas.Line (PicCanvas.ScaleLeft, PicCanvas.ScaleTop)-(PicCanvas.ScaleLeft + SCR, PicCanvas.ScaleTop + PicCanvas.ScaleHeight), , B
        For j = 0 To PicCanvas.ScaleTop + PicCanvas.ScaleHeight Step mvarScale
            'Draw big line
            PicCanvas.Line (PicCanvas.ScaleLeft, j)-(PicCanvas.ScaleLeft + SCR, j)
            'Print Value
            myValue = Round((j / mvarScale) * Div)
            PicCanvas.CurrentX = PicCanvas.ScaleLeft
            If j = 0 Then myValue = myValue + RSm
            PicCanvas.Print myValue
            'Draw small lines
            n = 0
            For i = j + mySmallScale To j + mvarScale Step mySmallScale
                n = n + 1
                If n = 5 Then
                    PicCanvas.Line (PicCanvas.ScaleLeft + (SCR / 2), i)-(PicCanvas.ScaleLeft + SCR, i)
                Else
                    PicCanvas.Line (PicCanvas.ScaleLeft + (SCR * 0.6), i)-(PicCanvas.ScaleLeft + SCR, i)
                End If
                If gScaleMode = 6 Then
                  PicCanvas.Line (PicCanvas.ScaleLeft, i)-(PicCanvas.ScaleLeft + SCR, i)
                  myValue = Round((i / mvarScale) * Div)
                  PicCanvas.CurrentX = PicCanvas.ScaleLeft
                  PicCanvas.Print myValue
                  For t = j To i Step mySmallScale / 10
                      PicCanvas.Line (PicCanvas.ScaleLeft + (SCR * 0.6), t)-(PicCanvas.ScaleLeft + SCR, t)
                   Next
                End If
            Next i
        Next j
        For j = 0 To PicCanvas.ScaleTop - 1000 Step -mvarScale
            'Draw big line
            PicCanvas.Line (PicCanvas.ScaleLeft, j)-(PicCanvas.ScaleLeft + SCR, j)
            'Print Value
            myValue = Round((j / mvarScale) * Div)
            PicCanvas.CurrentX = PicCanvas.ScaleLeft
            PicCanvas.Print myValue
            'Draw small lines
            n = 0
            For i = j + -mySmallScale To j + -mvarScale Step -mySmallScale
                n = n + 1
                If n = 5 Then
                    PicCanvas.Line (PicCanvas.ScaleLeft + (SCR / 2), i)-(PicCanvas.ScaleLeft + SCR, i)
                Else
                    PicCanvas.Line (PicCanvas.ScaleLeft + (SCR * 0.6), i)-(PicCanvas.ScaleLeft + SCR, i)
                End If
                If gScaleMode = 6 Then
                   PicCanvas.Line (PicCanvas.ScaleLeft, i)-(PicCanvas.ScaleLeft + SCR, i)
                   myValue = Round((i / mvarScale) * Div)
                   PicCanvas.CurrentX = PicCanvas.ScaleLeft
                   PicCanvas.Print myValue
                   For t = j To i Step -mySmallScale / 10
                      PicCanvas.Line (PicCanvas.ScaleLeft + (SCR * 0.6), t)-(PicCanvas.ScaleLeft + SCR, t)
                   Next
                End If
            Next i
        Next j
        
    End Select
    
    PicCanvas.Line (PicCanvas.ScaleLeft, PicCanvas.ScaleTop)-(PicCanvas.ScaleLeft + SCR, PicCanvas.ScaleTop + SCR), vbButtonFace, BF
    PicCanvas.Line (PicCanvas.ScaleLeft + 1 / gZoomFactor, PicCanvas.ScaleTop + 1 / gZoomFactor)-(PicCanvas.ScaleLeft + SCR - 1 / gZoomFactor, PicCanvas.ScaleTop + SCR - 1 / gZoomFactor), RGB(150, 150, 150), B
    PicCanvas.Line (PicCanvas.ScaleLeft + 1 / gZoomFactor, PicCanvas.ScaleTop + 1 / gZoomFactor)-(PicCanvas.ScaleLeft + SCR, PicCanvas.ScaleTop + SCR), vbWhite, B
    PicCanvas.Line (PicCanvas.ScaleLeft, PicCanvas.ScaleTop)-(PicCanvas.ScaleLeft + SCR, PicCanvas.ScaleTop + SCR), , B
    PicCanvas.Font.Size = Oldsize
    PicCanvas.ForeColor = OldFORECOLOR
End Sub

Private Sub icbDrawStyle_Click()
     icbDrawStyle.ToolTipText = icbDrawStyle.SelectedItem.Key
End Sub

Private Sub icbDrawWidth_Click()
    icbDrawWidth.ToolTipText = icbDrawWidth.SelectedItem.Key
End Sub

Private Sub List1_Click()
       If List1.ListIndex = -1 Then Exit Sub
       Text1.Text = m_OriginalPoints(List1.ListIndex + 1).X
       Text2.Text = m_OriginalPoints(List1.ListIndex + 1).Y
       LabelId.Caption = List1.ListIndex + 1
End Sub

Private Sub MeForm1_Hide()
       MeForm1.Visible = False
       m_ShowPenProperty = False
End Sub

Private Sub MeForm2_Hide()
     MeForm2.Visible = False
     m_ShowFillProperty = False
End Sub


Private Sub MeForm3_Hide()
     MeForm3.Visible = False
     m_ShowTranformProperty = False
End Sub

Private Sub MeForm4_Hide()
       MeForm4.Visible = False
       m_ShowObjectPoint = False
End Sub

Private Sub mnuAlternate_Click()
        If Not Obj Is Nothing Then
           Obj.FillMode = fALTERNATE
           Redraw
        End If
End Sub

Private Sub mnuClearTransform_Click()
   ClearTransform
End Sub

Private Sub mnuCurve_Click()
     Dim PointCoords() As PointAPI
     Dim PointType() As Byte, tmpType() As Byte
     Dim iCounter As Long, StartCounter As Long, EndCounter As Long
     Dim tx() As Single, ty() As Single, TPoint() As Byte
     Dim OldObj As vbdObject
     Dim txt As String, OldTxt As String, i As Long
     Dim xmin As Single, ymin As Single, xmax As Single, ymax As Single
          
     If Not Obj Is Nothing Then
     
       iCounter = 0
       Select Case Obj.TypeDraw
       Case dText
         ''Debug.Print Obj.Serialization
          Obj.Bound xmin, ymin, xmax, ymax
          PicCanvas.ForeColor = Obj.ForeColor
          PicCanvas.FillColor = Obj.FillColor
          BeginPath PicCanvas.hDC
          CenterText PicCanvas, (xmin + ((xmax - xmin) / 2)), (ymin + ((ymax - ymin) / 2)), _
                     Obj.TextDraw, Obj.Size, , -Obj.Angle * 10, Obj.Weight, _
                     Obj.Italic, Obj.Underline, Obj.Strikethrough, Obj.Charset, , , , , Obj.Name
          
          'CenterText PicCanvas, (xmin + ((xmax - xmin) / 2))  '/ gZoomFactor, (ymin + ((ymax - ymin) / 2))'/ gZoomFactor, _
                     Obj.TextDraw , Obj.size, , -Obj.Angle * 10, Obj.Weight, _
                     Obj.Italic, Obj.Underline, Obj.Strikethrough, Obj.Charset, , , , , Obj.Name
          EndPath PicCanvas.hDC
          
          iCounter = GetPathAPI(PicCanvas.hDC, ByVal 0&, ByVal 0&, 0)
           If (iCounter) Then
             ReDim PointCoords(iCounter - 1)
             ReDim PointType(iCounter - 1)
             'Get the path data from the DC
             Call GetPathAPI(PicCanvas.hDC, PointCoords(0), PointType(0), iCounter)
               StartCounter = 0
               EndCounter = iCounter - 1
          End If
          
       Case dEllipse, dFreePolygon, dPolygon, dPolyline, dScribble, dRectAngle, dPicture
            Obj.ReadPoint iCounter, tx, ty, tmpType
            ReDim PointCoords(1 To iCounter)
            ReDim PointType(1 To iCounter)
            StartCounter = 1
            EndCounter = iCounter
            PointType = tmpType
            i = 0
            For i = 1 To iCounter
                PointCoords(i).X = tx(i)
                PointCoords(i).Y = ty(i)
            Next
'            If Obj.TypeDraw = dRectAngle Then
'               ReDim Preserve PointCoords(1 To iCounter + 1)
'               ReDim Preserve PointType(1 To iCounter + 1)
'               StartCounter = 1
'               EndCounter = iCounter + 1
'               PointCoords(iCounter + 1).X = PointCoords(1).X
'               PointCoords(iCounter + 1).Y = PointCoords(1).Y
'               PointType(iCounter + 1) = PointType(1)
'               PointType(iCounter) = 3
'               iCounter = iCounter + 1
'            End If
       End Select
       
       If (iCounter) Then
             txt = txt & " DrawWidth(" + Trim(Str(Obj.DrawWidth)) + ")"
             txt = txt & " DrawStyle(" + Trim(Str(Obj.DrawStyle)) + ")"
             txt = txt & " ForeColor(" + Trim(Str(Obj.ForeColor)) + ")"
             txt = txt & " FillColor(" + Trim(Str(Obj.FillColor)) + ")"
             txt = txt & " FillColor2(" + Trim(Str(Obj.FillColor2)) + ")"
             txt = txt & " FillMode(" + Trim(Str(Obj.FillMode)) + ")"
             txt = txt & " Pattern(" + Obj.Pattern + ")"
             txt = txt & " Gradient(" + Trim(Str(Obj.Gradient)) + ")"
             If Not Obj.Picture Is Nothing Then
                txt = txt & " FillStyle(" + Trim(Str(10)) + ")"
             Else
             txt = txt & " FillStyle(" + Trim(Str(Obj.FillStyle)) + ")"
             End If
             txt = txt & " TypeDraw(" + Format$(dPolydraw) + ")"
             txt = txt & " TextDraw()"
             txt = txt & " CurrentX(" + Trim(Str(Obj.CurrentX)) + ")"
             txt = txt & " CurrentY(" + Trim(Str(Obj.CurrentY)) + ")"
             txt = txt & " TypeFill(" + Trim(Str(Obj.TypeFill)) + ")"
             txt = txt & " ObjLock(" + Trim(Str(Obj.ObjLock)) + ")"
             txt = txt & " Blend(" + Trim(Str(Obj.Blend)) + ")"
             txt = txt & " Shade(False)"
             txt = txt & " AlingText(0)"
             txt = txt & " Bold(0)"
             txt = txt & " Charset(0)"
             txt = txt & " Italic(0)"
             txt = txt & " Name()"
             txt = txt & " Size(0)"
             txt = txt & " Strikethrough(0)"
             txt = txt & " Underline(0)"
             txt = txt & " Weight(400)"
             txt = txt & " Angle(" + Trim(Str(Obj.Angle)) + ")"
             txt = txt & vbCr & "Transformation(1 0 0 0 1 0 0 0 1 )"
             txt = txt & " IsClosed(True)"
             txt = txt & " NumPoints(" & Format$(iCounter) & ")"
    
             For i = StartCounter To EndCounter
                 txt = txt & vbCrLf & "    X(" & Format$(PointCoords(i).X) & ")"
                 txt = txt & " Y(" & Format$(PointCoords(i).Y) & ")"
                 txt = txt & " P(" & Format$(PointType(i)) & ")"
             Next i
             If Not Obj.Picture Is Nothing Then
                SavePicture Obj.Picture, Environment(s_TEMP) + "\Temp.bmp"
                NewImage = OpenFileEnc(Environment(s_TEMP) + "\Temp.bmp")
                txt = txt & vbCrLf + " Image (" + vbCrLf + NewImage + ")"
                If FileExists(Environment(s_TEMP) + "\Temp.bmp") Then Kill Environment(s_TEMP) + "\Temp.bmp"
             End If
             txt = "PolyDraw(PolyDraw(" & txt & "))"
             ObjectDelete
             OldTxt = Clipboard.GetText
             Clipboard.SetText txt
             PasteObject
             Clipboard.SetText OldTxt
       End If
       End If
End Sub

'Only Polygon
Private Sub mnueditPoints_Click()

Dim oType() As Byte, oX() As Single, oY() As Single, oNumPoints As Long, oPointCoods() As PointAPI, sPointCoods() As PointAPI
Dim PointType() As Byte, tx() As Single, ty() As Single
Dim iCounter As Long, xmin As Single, ymin As Single, xmax As Single, ymax As Single, aa As Long

       Obj.Bound xmin, ymin, xmax, ymax
       Obj.ReadTrPoint oNumPoints, oX(), oY(), oType()
       
       F = InputBox("Number of points (3-20)", "Polygon")
       If F <> "" Then
        If Val(F) >= 3 And Val(F) <= 20 Then
            oNumPoints = Val(F)
            oPointCoods = CreatePolygon(xmin + (xmax - xmin) / 2, ymin + (ymax - ymin) / 2, oX(1), oY(1), oNumPoints)
            sPointCoods = CreatePolygon(xmin + (xmax - xmin) / 2, ymin + (ymax - ymin) / 2, oX(2), oY(2), oNumPoints)
            
            oNumPoints = oNumPoints * 2 + 1
            ReDim tx(1 To oNumPoints)
            ReDim ty(1 To oNumPoints)
            ReDim PointType(1 To oNumPoints)
            aa = 0
            For iCounter = 1 To oNumPoints - 1 Step 2
               aa = aa + 1
               tx(iCounter) = oPointCoods(aa).X
               ty(iCounter) = oPointCoods(aa).Y
               PointType(iCounter) = 2
            Next
            aa = 0
            For iCounter = 2 To oNumPoints - 1 Step 2
               aa = aa + 1
               tx(iCounter) = sPointCoods(aa).X
               ty(iCounter) = sPointCoods(aa).Y
               PointType(iCounter) = 2
            Next
            tx(oNumPoints) = tx(1)
            ty(oNumPoints) = ty(1)
            PointType(1) = 6
            PointType(oNumPoints) = 3
            Obj.NewPoint oNumPoints, tx, ty, PointType
         End If
       End If
End Sub

Private Sub mnuFillMode_Click()
     mnutransform_Click
End Sub

Private Sub mnuLock_Click()
       ObjectLock True
End Sub

Private Sub mnuProperty_Click()
   Dim Msg As String
   Dim X1 As Single, X2 As Single, Y1 As Single, Y2 As Single
      If Not Obj Is Nothing Then
         Obj.Bound X1, Y1, X2, Y2
         Msg = "StartX:" + Str(X1) + " - StartY:" + Str(Y1) + vbCr
         Msg = Msg + "EndX:" + Str(X1) + " - EndY:" + Str(Y1) + vbCr
         Msg = Msg + Obj.Info + vbCr
         Msg = Msg + "FillColor1:" + Str(Obj.FillColor) + vbCr
         Msg = Msg + "FillColor2:" + Str(Obj.FillColor2) + vbCr
         Msg = Msg + "ForeColor:" + Str(Obj.ForeColor) + vbCr
         Msg = Msg + "FillMode:" + Str(Obj.FillMode) + vbCr
         Msg = Msg + "FillStyle:" + Str(Obj.FillStyle) + vbCr
         Msg = Msg + "Gradient:" + Str(Obj.Gradient) + vbCr
         Msg = Msg + "Blend:" + Str(Obj.Blend) + vbCr
         Msg = Msg + "Pattern:" + Obj.Pattern
         MsgBox Msg, vbInformation
      End If
End Sub

Private Sub mnutransform_Click()
      If Not Obj Is Nothing Then
         Select Case Obj.FillMode
         Case 1
            mnuAlternate.Checked = True
            mnuWinding.Checked = False
         Case 2
            mnuAlternate.Checked = False
            mnuWinding.Checked = True
         Case Else
            mnuAlternate.Checked = False
            mnuWinding.Checked = False
         End Select
         
         If Obj.ObjLock = True Then
            mnuLock.Enabled = False
            mnuUnlock.Enabled = True
         Else
            mnuLock.Enabled = True
            mnuUnlock.Enabled = False
         End If
'         Obj.TypeDraw = dText Or
         If Obj.TypeDraw = dEllipse _
            Or Obj.TypeDraw = dFreePolygon Or Obj.TypeDraw = dPolygon _
            Or Obj.TypeDraw = dPolyline Or Obj.TypeDraw = dRectAngle _
            Or Obj.TypeDraw = dTextFrame Then
            mnuCurve.Enabled = True
         Else
            mnuCurve.Enabled = False
         End If
         
         mnuProperty.Enabled = True
      Else
         mnuCurve.Enabled = False
         mnuAlternate.Checked = False
         mnuWinding.Checked = False
         mnuLock.Enabled = False
         mnuUnlock.Enabled = False
         mnuProperty.Enabled = False
      End If
End Sub

Private Sub mnuUnlock_Click()
      ObjectLock False
       
End Sub

Private Sub mnuWinding_Click()
      If Not Obj Is Nothing Then
           Obj.FillMode = fWINDING
           Redraw
      End If
End Sub

Private Sub picCanvas_DblClick()
    EditText
End Sub

Private Sub picCanvas_KeyUp(KeyCode As Integer, Shift As Integer)
      Dim Msg As String
      Dim X_min As Single
      Dim Y_min As Single
      If Obj Is Nothing Then Exit Sub
      
      Select Case KeyCode
      Case vbKeyLeft '37 LEFT ARROW key
          X_min = -1
          Y_min = 0
      Case vbKeyUp '38 UP ARROW key
          Y_min = -1
          X_min = 0
      Case vbKeyRight '39 RIGHT ARROW key
          X_min = 1
          Y_min = 0
      Case vbKeyDown '40
          Y_min = 1
          X_min = 0
      End Select

      'TransformPoint X_min / m_ZoomFactor, Y_min / m_ZoomFactor
      TransformPoint X_min, Y_min
End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
        
         RaiseEvent MouseDown(Button, Shift, X, Y)
         
'        'check if click in ruler
'        Dim SCR As Long, TW As Single
'        SCR = Round((DataYmax - DataYmin) / (Wymax - Wymin), 3)
'        SCR = 50 / SCR
'        TW = PicCanvas.TextWidth("Pix")
'        If TW > SCR Then
'           SCR = TW * 2
'        ElseIf TW * 2 < SCR Then
'           SCR = TW * 2
'        End If
'        If X <= PicCanvas.ScaleLeft + SCR Then Stop
'        If Y <= PicCanvas.ScaleTop + SCR Then Stop
        
         If Not Obj Is Nothing And m_EditPoint = False Then

         End If
         
         Select Case DrawingMode
         Case MODE_ReadFill, MODE_ReadPen
             Call ReadColorProperty(X, Y)
             
         Case MODE_EDITOBJ
             Call MouseDown(Button, Shift, X, Y)
             
         Case MODE_EDITPOINT
              Call MouseDownPoint(Button, Shift, X, Y)
              
         Case MODE_START_ZOOM
           ' Start a zooming rubberband hex.
            DrawingMode = MODE_ZOOMING
        
            OldMode = PicCanvas.DrawMode
            PicCanvas.DrawMode = vbInvert
            
            zStartX = X
            zStartY = Y
            zLastX = X
            zLastY = Y
            PicCanvas.Line (zStartX, zStartY)-(zLastX, zLastY), , B
            
         Case MODE_PANNING
            If Button = vbLeftButton And Shift = 0 Then
               zStartX = X
               zStartY = Y
               MouseIcon 22
            End If
           
         End Select
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim XM As Single, yM As Single, ZFactor As Single, SCR As Single
            
    If CrossMouse = True Then
       RedrawCross X, Y
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
       
    ScrollMouse = GetScrollMovement(PicCanvas.hWnd)
    
    If ScrollType = 0 Then
       CenterZoomX = X
       CenterZoomY = Y
       If ScrollMouse = 1 Then 'Zoom in
          ZFactor = Round(gZoomFactor, 2) + 0.5
          If ZFactor > 10 Then ZFactor = 10
          SetScaleFull False
          SetScaleFactor ZFactor
       ElseIf ScrollMouse = -1 Then 'Zoom out
          ZFactor = Round(gZoomFactor, 2) - 0.5
          If ZFactor < 0.1 Then ZFactor = 0.1
          SetScaleFull False
          SetScaleFactor ZFactor
       Else 'The event was fired for a proper MouseMove
         '<<DEAL WITH THE OTHER MOUSEMOVE stuff here>>
       End If
       RaiseEvent ZoomChange
    Else
       If VScroll1.Visible = True Then
          If ScrollMouse = 1 Then 'Scroll up
             VScroll1.Value = VScroll1.Value + 50
          ElseIf ScrollMouse = -1 Then 'Scroll down
              VScroll1.Value = VScroll1.Value - 50
          Else 'The event was fired for a proper MouseMove
               '<<DEAL WITH THE OTHER MOUSEMOVE stuff here>>
          End If
       End If
    End If
     
    If Button = vbLeftButton And Shift = 0 Then
        Select Case DrawingMode
        Case MODE_EDITOBJ
            Call MouseMove(Button, Shift, X, Y)
        Case MODE_EDITPOINT
            Call MouseMovePoint(Button, Shift, X, Y)
        Case MODE_ZOOMING
            ' Erase the old hex.
            PicCanvas.Line (zStartX, zStartY)-(zLastX, zLastY), , B
    
            ' Draw the new hex.
            zLastX = X
            zLastY = Y
            PicCanvas.Line (zStartX, zStartY)-(zLastX, zLastY), , B
        
        Case MODE_PANNING
            With PicCanvas
                zLastY = -((Y - zStartY))
                zLastX = -((X - zStartX))
                zLastY = -zLastY
                zLastX = -zLastX
            End With
        End Select
    
    Else
    
       Select Case DrawingMode
       Case MODE_EDITOBJ
            Call MouseMove(Button, Shift, X, Y)
       Case MODE_EDITPOINT
          Dim mStep As Integer
          mStep = IIf(GAP / gZoomFactor > 0, GAP / gZoomFactor, 3)
          PicCanvas.ToolTipText = ""
           
          For i = 1 To m_NumPoints
            If X >= m_OriginalPoints(i).X - mStep And X <= m_OriginalPoints(i).X + mStep And _
               Y >= m_OriginalPoints(i).Y - mStep And Y <= m_OriginalPoints(i).Y + mStep Then
               'm_SelectPoint = i
               PicCanvas.ToolTipText = "Select Point:" + Str(i)
               RaiseEvent MsgControl(PicCanvas.ToolTipText)
               Exit For
             End If
           Next
        Case Else
            PicCanvas.ToolTipText = ""
       End Select
    End If
    
End Sub


Private Sub PicCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
    RaiseEvent MouseUp(Button, Shift, X, Y)
      
    If Button = vbLeftButton And Shift = 0 Then

    Select Case DrawingMode
    Case MODE_EDITOBJ
         Call MouseUp(Button, Shift, X, Y)
         
    Case MODE_EDITPOINT
         Call MouseUpPoint(Button, Shift, X, Y)
         
    Case MODE_ZOOMING
         
        ' Erase the old hex.
         PicCanvas.Line (zStartX, zStartY)-(zLastX, zLastY), , B
         zLastX = X
         zLastY = Y
    
        ' We're done drawing for this rubberband hex.
         PicCanvas.DrawMode = OldMode
         PicCanvas.MousePointer = vbDefault
         If zStartX + 10 > zLastX Or zStartY + 10 > zLastY Then
            GoTo mu1
         End If
         ' Set the new world window bounds.
         If zStartX > zLastX Then
            Wxmin = zLastX
            Wxmax = zStartX
         Else
            Wxmin = zStartX
            Wxmax = zLastX
         End If
        If zStartY > zLastY Then
           Wymin = zLastY
           Wymax = zStartY
        Else
           Wymin = zStartY
           Wymax = zLastY
        End If
        
        ' Set the new world window bounds.
         SetWorldWindow
         RaiseEvent ZoomChange
         Redraw
         
mu1:
         DrawingMode = OldDrawingMode 'MODE_NONE
         If DrawingMode = 0 Then DrawingMode = 1
         MouseIcon DrawingMode
         SelectTool DrawingMode
         
    Case MODE_PANNING
         If VScroll1.Visible Then
            If VScroll1.Value - zLastY * 10 >= VScroll1.Min And VScroll1.Value - zLastY * 10 <= VScroll1.Max Then
              'VScroll1.Value = VScroll1.Max
              VScroll1.Value = VScroll1.Value - zLastY * 10
           ElseIf VScroll1.Value - zLastY * 10 < VScroll1.Min Then
              VScroll1.Value = VScroll1.Min
            ElseIf VScroll1.Value - zLastY * 10 > VScroll1.Max Then
              VScroll1.Value = VScroll1.Max
           End If
         End If
         If HScroll1.Visible Then
           If HScroll1.Value - zLastX * 10 >= HScroll1.Min And HScroll1.Value - zLastX * 10 <= HScroll1.Max Then
              HScroll1.Value = HScroll1.Value - zLastX * 10
           ElseIf HScroll1.Value - zLastX * 10 > HScroll1.Max Then
              HScroll1.Value = HScroll1.Max
           ElseIf HScroll1.Value - zLastX * 10 < HScroll1.Min Then
              HScroll1.Value = HScroll1.Min
           End If
         End If
       ' Set the new world window bounds.
        SetWorldWindow
        Redraw
        DrawingMode = OldDrawingMode
        If DrawingMode = 0 Then DrawingMode = 1
        SelectTool DrawingMode
        MouseIcon DrawingMode
        
    End Select
    End If
      
End Sub

Private Sub ClearBox()
    XminBox = 0
    XmaxBox = 0
    YminBox = 0
    YmaxBox = 0
End Sub

Private Sub picCanvas_Paint()
   Dim mDC As Long, tmpBmp As Long
    If mLockControl Then Exit Sub
    
    If m_TheScene Is Nothing Then Exit Sub
    
    PicCanvas.Cls
    PicCanvas.DrawStyle = 0
    ReDrawPage
    PicCanvas.DrawMode = 13
    LockWindowUpdate UserControl.hWnd
    m_TheScene.Draw PicCanvas
    If m_DrawRuler = True Then
       ReDrawRuler 1
       ReDrawRuler 2
    End If
    LockWindowUpdate False
   
End Sub

Sub ViewMenu()
'Obj.TypeDraw = dText Or
    If Obj.TypeDraw = dEllipse _
       Or Obj.TypeDraw = dFreePolygon _
       Or Obj.TypeDraw = dPolyline Or Obj.TypeDraw = dRectAngle _
       Or Obj.TypeDraw = dTextFrame Then
       sep1.Visible = True
       mnuCurve.Visible = True
       mnuFillMode.Enabled = False
    Else
       mnuCurve.Visible = False
       mnuFillMode.Enabled = True
    End If
    
    If Obj.TypeDraw = dPolygon Then
       sepEditPoints.Visible = True
       mnueditPoints.Visible = True
    Else
       sepEditPoints.Visible = False
       mnueditPoints.Visible = False
    End If
   
    PopupMenu mnutransform
End Sub

Private Sub PicPenColor_DblClick()
    cmdSysColorsPen_Click
End Sub


Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
             ObjectDelete
    End Select
End Sub

Private Sub UserControl_Resize()
    Dim TW As Long, tH As Long, nW As Long, toZoom As Boolean
    Dim pcW As Long, pcH As Long, pcL As Long, pcT As Long
    
   If mLockControl = True Then Exit Sub
        
    Dim X As Single
    Dim Y As Single
    Dim wid As Single
    Dim hgt As Single
   
    ' Fit the viewport to the window.
    X = 0
    Y = 0
    wid = UserControl.ScaleWidth
    hgt = UserControl.ScaleHeight
    
    PicCanvas.Move X, Y, wid, hgt
    ''Debug.Print "Resize", X, Y, wid, hgt
    VAspect = hgt / wid
    
    ' Place the scroll bars next to the viewport.
    X = PicCanvas.Left + PicCanvas.Width - VScroll1.Width
    Y = PicCanvas.Top
    wid = VScroll1.Width
    hgt = PicCanvas.Height - HScroll1.Height
    VScroll1.Move X, Y, wid, hgt
    
    X = PicCanvas.Left
    Y = PicCanvas.Top + PicCanvas.Height - HScroll1.Height
    wid = PicCanvas.Width - VScroll1.Width
    hgt = HScroll1.Height
    HScroll1.Move X, Y, wid, hgt
    
    ComCorner.Move VScroll1.Left, HScroll1.Top, VScroll1.Width, HScroll1.Height
    
    ' Start at full scale.
    SetScaleFull
          
    If m_ShowPenProperty = True Then MeForm1.Visible = True
    If m_ShowFillProperty = True Then MeForm2.Visible = True
    If m_ShowTranformProperty = True Then MeForm3.Visible = True
    If mLockControl = False Then PicCanvas.Visible = True
    If m_ShowObjectPoint = True Then MeForm4.Visible = True
    RaiseEvent ZoomChange
   ' Redraw
    
End Sub

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = PicCanvas.Image
End Property

Public Property Set Image(ByVal New_Image As Picture)
    Set m_Image = New_Image
    PropertyChanged "Image"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set m_Image = LoadPicture("")
    
    m_CanvasWidth = m_def_CanvasWidth
    m_CanvasHeight = m_def_CanvasHeight
    m_ShowCanvasSize = m_def_ShowCanvasSize
    m_ForeColor = m_def_ForeColor
    m_DrawWidth = m_def_DrawWidth
    m_DrawStyle = m_def_DrawStyle
    m_FillStyle = m_def_FillStyle
    m_FillColor = m_def_FillColor
    m_FileName = m_def_FileName
    m_FileTitle = m_def_FileTitle
    m_ShowPenProperty = m_def_ShowPenProperty
    m_ShowFillProperty = m_def_ShowFillProperty
    m_ShowTranformProperty = m_def_ShowFillProperty
    m_FillColor2 = m_def_FillColor2
    m_Pattern = m_def_Pattern
    m_TypeGradient = m_def_TypeGradient
    m_Blend = m_def_Blend
    m_BackImage = m_def_BackImage
    m_LockObject = m_def_LockObject
    m_hDC = m_def_hDC
    Set m_ObjPicture = LoadPicture("")
    m_EditPoint = m_def_EditPoint
    m_CrossMouse = m_def_CrossMouse
    m_DrawRuler = m_def_DrawRuler
    m_BackColor = m_def_BackColor
    
    NewDraw
    
    m_ShowObjectPoint = m_def_ShowObjectPoint
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    PicCanvas.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Set m_Image = PropBag.ReadProperty("Image", Nothing)
    
    m_CanvasWidth = PropBag.ReadProperty("CanvasWidth", m_def_CanvasWidth)
    m_CanvasHeight = PropBag.ReadProperty("CanvasHeight", m_def_CanvasHeight)
    m_ShowCanvasSize = PropBag.ReadProperty("ShowCanvasSize", m_def_ShowCanvasSize)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_DrawWidth = PropBag.ReadProperty("DrawWidth", m_def_DrawWidth)
    m_DrawStyle = PropBag.ReadProperty("DrawStyle", m_def_DrawStyle)
    m_FillStyle = PropBag.ReadProperty("FillStyle", m_def_FillStyle)
    m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
    m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
    m_FileTitle = PropBag.ReadProperty("FileTitle", m_def_FileTitle)
    m_ShowPenProperty = PropBag.ReadProperty("ShowPenProperty", m_def_ShowPenProperty)
    m_ShowFillProperty = PropBag.ReadProperty("ShowFillProperty", m_def_ShowFillProperty)
    m_ShowTranformProperty = PropBag.ReadProperty("ShowTranformProperty", m_def_ShowTranformProperty)
    m_FillColor2 = PropBag.ReadProperty("FillColor2", m_def_FillColor2)
    m_Pattern = PropBag.ReadProperty("Pattern", m_def_Pattern)
    m_TypeGradient = PropBag.ReadProperty("Gradient", m_def_TypeGradient)
    m_Blend = PropBag.ReadProperty("Blend", m_def_Blend)
    m_BackImage = PropBag.ReadProperty("BackImage", m_def_BackImage)
    m_LockObject = PropBag.ReadProperty("LockObject", m_def_LockObject)
    m_hDC = PropBag.ReadProperty("hDC", m_def_hDC)
    Set m_ObjPicture = PropBag.ReadProperty("ObjPicture", Nothing)
    m_EditPoint = PropBag.ReadProperty("EditPoint", m_def_EditPoint)
    m_CrossMouse = PropBag.ReadProperty("CrossMouse", m_def_CrossMouse)
    m_DrawRuler = PropBag.ReadProperty("DrawRuler", m_def_DrawRuler)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)

    MeForm1.Caption = "Pen"
    MeForm1.Alignment = 0
    MeForm2.Caption = "Fill"
    MeForm2.Alignment = 0
    MeForm3.Caption = "Transform"
    MeForm3.Alignment = 0
    MeForm4.Caption = "Object Point"
    MeForm4.Alignment = 0
    MeForm1.BackColor = &HF1E2DC
    MeForm2.BackColor = &HF1E2DC
    MeForm3.BackColor = &HF1E2DC
    MeForm4.BackColor = &HF1E2DC
    CtlFill1.BackColor = &HF1E2DC
    CtrTranform1.BackColor = &HF1E2DC
    MeForm1.Move UserControl.ScaleWidth - MeForm1.Width + 1 - VScroll1.Width, 0
    MeForm1.Height = 190 'UserControl.ScaleHeight / 3
    MeForm2.Move UserControl.ScaleWidth - MeForm1.Width + 1 - VScroll1.Width, MeForm1.Height
    MeForm2.Height = 210 'UserControl.ScaleHeight / 3
    MeForm3.Move UserControl.ScaleWidth - MeForm1.Width + 1 - VScroll1.Width, MeForm1.Height + MeForm2.Height
    MeForm3.Height = 165 'UserControl.ScaleHeight / 3
    MeForm4.Move UserControl.ScaleWidth - MeForm1.Width + 1 - VScroll1.Width, 0
      
    DrawPen
    InitPage
    NewDraw
       
    m_ShowObjectPoint = PropBag.ReadProperty("ShowObjectPoint", m_def_ShowObjectPoint)
End Sub

Private Sub UserControl_Show()
   If IsDebugMode = False Then
      AddScrollness PicCanvas.hWnd
   End If
End Sub

Private Sub UserControl_Terminate()
   If IsDebugMode = False Then
      RemoveScrollness PicCanvas.hWnd
   End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("BackColor", PicCanvas.BackColor, &H80000005)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("Image", m_Image, Nothing)
    
    Call PropBag.WriteProperty("CanvasWidth", m_CanvasWidth, m_def_CanvasWidth)
    Call PropBag.WriteProperty("CanvasHeight", m_CanvasHeight, m_def_CanvasHeight)
    Call PropBag.WriteProperty("ShowCanvasSize", m_ShowCanvasSize, m_def_ShowCanvasSize)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("DrawWidth", m_DrawWidth, m_def_DrawWidth)
    Call PropBag.WriteProperty("DrawStyle", m_DrawStyle, m_def_DrawStyle)
    Call PropBag.WriteProperty("FillStyle", m_FillStyle, m_def_FillStyle)
    Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
    Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
    Call PropBag.WriteProperty("FileTitle", m_FileTitle, m_def_FileTitle)
    Call PropBag.WriteProperty("ShowPenProperty", m_ShowPenProperty, m_def_ShowPenProperty)
    Call PropBag.WriteProperty("ShowFillProperty", m_ShowFillProperty, m_def_ShowFillProperty)
    Call PropBag.WriteProperty("ShowTranformProperty", m_ShowTranformProperty, m_def_ShowTranformProperty)
    Call PropBag.WriteProperty("FillColor2", m_FillColor2, m_def_FillColor2)
    Call PropBag.WriteProperty("Pattern", m_Pattern, m_def_Pattern)
    Call PropBag.WriteProperty("Gradient", m_TypeGradient, m_def_TypeGradient)
    Call PropBag.WriteProperty("Blend", m_Blend, m_def_Blend)
    Call PropBag.WriteProperty("BackImage", m_BackImage, m_def_BackImage)
    Call PropBag.WriteProperty("LockObject", m_LockObject, m_def_LockObject)
    Call PropBag.WriteProperty("hDC", m_hDC, m_def_hDC)
    Call PropBag.WriteProperty("ObjPicture", m_ObjPicture, Nothing)
    Call PropBag.WriteProperty("EditPoint", m_EditPoint, m_def_EditPoint)
    Call PropBag.WriteProperty("CrossMouse", m_CrossMouse, m_def_CrossMouse)
    Call PropBag.WriteProperty("DrawRuler", m_DrawRuler, m_def_DrawRuler)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ShowObjectPoint", m_ShowObjectPoint, m_def_ShowObjectPoint)
End Sub

Public Property Get CanvasWidth() As Long
    CanvasWidth = m_CanvasWidth
End Property

Public Property Let CanvasWidth(ByVal New_CanvasWidth As Long)
    m_CanvasWidth = New_CanvasWidth
    PropertyChanged "CanvasWidth"
End Property

Public Property Get CanvasHeight() As Long
    CanvasHeight = m_CanvasHeight
End Property

Public Property Let CanvasHeight(ByVal New_CanvasHeight As Long)
    m_CanvasHeight = New_CanvasHeight
    PropertyChanged "CanvasHeight"
End Property

Private Sub VScroll1_Change()

    If IgnoreSbarChange Then Exit Sub
    VScrollBarChanged
End Sub

Public Property Get ShowCanvasSize() As Boolean
    ShowCanvasSize = m_ShowCanvasSize
End Property

Public Property Let ShowCanvasSize(ByVal New_ShowCanvasSize As Boolean)
    m_ShowCanvasSize = New_ShowCanvasSize
    PropertyChanged "ShowCanvasSize"
End Property

Public Sub Redraw()
    picCanvas_Paint
End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    ChangeForeColor m_ForeColor
    RaiseEvent ColorSelected(2, m_ForeColor)
End Property

Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns/sets the line width for output from graphics methods."
    DrawWidth = m_DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    m_DrawWidth = New_DrawWidth
    PropertyChanged "DrawWidth"
    ChangeDrawWidth m_DrawWidth
End Property

Public Property Get DrawStyle() As Integer
Attribute DrawStyle.VB_Description = "Determines the line style for output from graphics methods."
    DrawStyle = m_DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As Integer)
    m_DrawStyle = New_DrawStyle
    PropertyChanged "DrawStyle"
    ChangeDrawstyle m_DrawStyle
End Property

Public Property Get FillStyle() As Integer
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
    FillStyle = m_FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As Integer)
    m_FillStyle = New_FillStyle
    PropertyChanged "FillStyle"
    ChangeFillstyle m_FillStyle
End Property

Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColor = m_FillColor
End Property

Public Property Let FillColor(ByVal new_FillColor As OLE_COLOR)
    m_FillColor = new_FillColor
    PropertyChanged "FillColor"
    ChangeFillColor 1, m_FillColor
    RaiseEvent ColorSelected(1, m_FillColor)
End Property

Public Sub SelectTool(ByVal Key As Integer)
    Dim Msg As String
    Dim new_pgon As vbdPolygon
    Dim new_plgon As VbPolygon
    Dim new_line As vbdLine
    Dim new_text As VbText
    Dim new_Scribble As vbdScribble
    Dim new_Ellipse As vbdEllipse
    Dim new_curve As vbdPolygon 'vbdCurve
    Dim m_ToolKey As Integer
    
    ' Free any previously started object.
    Set m_NewObject = Nothing
        
    m_DrawingObject = True
   ' m_EditPoint = False
    
    If Key <> 1 And Key <> 2 Then
        DeselectAll
    End If
    
    ' Create the new object.
    m_ToolKey = Key
    
    Select Case m_ToolKey
        Case 1 '"Arrow"
            MouseIcon 1
            Msg = "Select object"
            m_DrawingObject = False
            DrawingMode = MODE_EDITOBJ
            If Not Obj Is Nothing Then
               Obj.Selected = True
               Obj.EditPoint = False
               Redraw
            End If
           m_NumPoints = 0
            Erase MX
            Erase mY
            Erase m_TypePoint
            Erase m_OriginalPoints
            
        Case 2 '"Point"
            
            If Not Obj Is Nothing Then
                If Obj.ObjLock = True Then Exit Sub
                    If Obj.TypeDraw = [dPicture] Or _
                       Obj.TypeDraw = [dEllipse] Then
                        Exit Sub
                    End If
            End If
            DrawingMode = MODE_EDITPOINT
            MouseIcon 2
            SaveSnapshot
          
            If Not Obj Is Nothing Then
                Obj.Selected = False
                Obj.EditPoint = True
                Obj.ReadTrPoint m_NumPoints, MX(), mY(), m_TypePoint
                
                Obj.NewPoint m_NumPoints, MX, mY, m_TypePoint
                ReDim m_OriginalPoints(1 To m_NumPoints)
                For r = 1 To m_NumPoints
                    m_OriginalPoints(r).X = MX(r)
                    m_OriginalPoints(r).Y = mY(r)
                Next
                Redraw
                DrawPoint
           End If

        Case 3 '"Polyline"
             DrawingMode = MODE_POLYLINE
            Set m_NewObject = New vbdPolygon
            Set new_pgon = m_NewObject
            new_pgon.IsClosed = False
            new_pgon.m_DrawStyle = 0
            new_pgon.m_DrawWidth = 1
            new_pgon.m_FillColor = RGB(255, 255, 255)
            new_pgon.m_FillStyle = 1
            new_pgon.m_ForeColor = RGB(0, 0, 0)
            new_pgon.m_TypeDraw = dPolyline
            new_pgon.m_Shade = False
            MouseIcon 3
            Msg = "Draw Line or Polyline"
            
        Case 4 '"FreePolygon"
             DrawingMode = MODE_POLYLINE
            Set m_NewObject = New vbdPolygon
            Set new_pgon = m_NewObject
            new_pgon.IsClosed = True
            new_pgon.m_DrawStyle = 0
            new_pgon.m_DrawWidth = 1
            new_pgon.m_FillColor = RGB(255, 255, 255)
            new_pgon.m_FillStyle = 1
            new_pgon.m_ForeColor = RGB(0, 0, 0)
            new_pgon.m_TypeDraw = dFreePolygon
            new_pgon.m_Shade = False
            MouseIcon 4
            Msg = "Draw free Polygon"
        
        Case 5 'Free line "Scribble"
             DrawingMode = MODE_Scribble
            Set m_NewObject = New vbdScribble
            Set new_Scribble = m_NewObject
            new_Scribble.m_DrawStyle = 0
            new_Scribble.m_DrawWidth = 1
            new_Scribble.m_FillColor = RGB(255, 255, 255)
            new_Scribble.m_FillStyle = 1
            new_Scribble.m_ForeColor = RGB(0, 0, 0)
            new_Scribble.m_TypeDraw = dScribble
            new_Scribble.IsClosed = False
            new_Scribble.m_Shade = False
            MouseIcon 5
            Msg = "Draw free line"
            
        Case 6 ' "Calligraphic"
             DrawingMode = MODE_Scribble
            Set m_NewObject = New vbdScribble
            Set new_Scribble = m_NewObject
            new_Scribble.m_DrawStyle = 0
            new_Scribble.m_DrawWidth = 1
            
            new_Scribble.m_FillColor = FillColor ' 'RGB(255, 255, 255)
          '  new_Scribble.m_FillStyle = 1
            new_Scribble.m_FillMode = fWINDING
            new_Scribble.m_ForeColor = ForeColor 'RGB(0, 0, 0)
            new_Scribble.m_TypeDraw = dCalligraphic
            new_Scribble.IsClosed = False
            new_Scribble.m_Shade = False
            new_Scribble.m_Blend = 127
            MouseIcon 6
            Msg = "Draw free line close"
              
        Case 7 '"Curve"
            Set m_NewObject = New vbdPolygon ' vbdCurve
            Set new_curve = m_NewObject
            new_curve.IsClosed = False
            new_curve.m_DrawStyle = 0
            new_curve.m_DrawWidth = 1
            new_curve.m_FillColor = RGB(255, 255, 255)
            new_curve.m_FillStyle = 1
            new_curve.m_ForeColor = RGB(0, 0, 0)
            new_curve.m_TypeDraw = dCurve
            new_curve.m_Shade = False
            MouseIcon 7
            Msg = "Select 2 point to draw curve"
            
        Case 8 '"RectAngle"
            DrawingMode = MODE_RECTANGLE
            Set m_NewObject = New vbdLine
            Set new_line = m_NewObject
            new_line.IsBox = True
            new_line.m_DrawStyle = 0
            new_line.m_DrawWidth = 1
            new_line.m_FillColor = RGB(255, 255, 255)
            new_line.m_FillStyle = 1
            new_line.m_ForeColor = RGB(0, 0, 0)
            new_line.m_TypeDraw = dRectAngle
            new_line.m_Shade = False
            MouseIcon 8
            Msg = "Press and Hold ''Ctrl'' Button to make a Cube"
            
        Case 9 '"Polygon"
             DrawingMode = MODE_POLYGON
            Set m_NewObject = New VbPolygon
            Set new_plgon = m_NewObject
            new_plgon.IsBox = True
            new_plgon.m_DrawStyle = 0
            new_plgon.m_DrawWidth = 1
            new_plgon.m_FillColor = RGB(255, 255, 255)
            new_plgon.m_FillStyle = 1
            new_plgon.m_ForeColor = RGB(0, 0, 0)
            new_plgon.m_TypeDraw = dPolygon
            new_plgon.m_Shade = False
            MouseIcon 9
            Msg = "Press and Hold ''Ctrl'' Button to make a Polygon"
        
        Case 10 '"Ellipse"
            DrawingMode = MODE_ELLIPSE
            Set m_NewObject = New vbdEllipse
            Set new_Ellipse = m_NewObject
            new_Ellipse.m_DrawStyle = 0
            new_Ellipse.m_DrawWidth = 1
            new_Ellipse.m_FillColor = RGB(255, 255, 255)
            new_Ellipse.m_FillStyle = 1
            new_Ellipse.m_ForeColor = RGB(0, 0, 0)
            new_Ellipse.m_TypeDraw = dEllipse
            new_Ellipse.m_Shade = False
            MouseIcon 10
            Msg = "Press and Hold ''Ctrl'' Button to make a Circle"
            
        Case 11 '"Text"
            DrawingMode = MODE_TEXT
            Set m_NewObject = New VbText
            Set new_text = m_NewObject
            new_text.IsBox = True
            new_text.m_DrawStyle = 0
            new_text.m_DrawWidth = 1
            new_text.m_FillColor = RGB(255, 255, 255)
            new_text.m_FillStyle = 1
            new_text.m_ForeColor = RGB(0, 0, 0)
            new_text.m_TypeDraw = dText
            new_text.m_Shade = False
            MouseIcon 11
            Msg = "Select position for text"
            
'        Case 11 '"TextArt"
'            DrawingMode = MODE_TEXTFRAME
'            Set m_NewObject = New VbText
'            Set new_text = m_NewObject
'            new_text.IsBox = True
'            new_text.m_DrawStyle = 0
'            new_text.m_DrawWidth = 1
'            new_text.m_FillColor = RGB(255, 255, 255)
'            new_text.m_FillStyle = 1
'            new_text.m_ForeColor = RGB(0, 0, 0)
'            new_text.m_TypeDraw = dTextFrame
'            new_text.m_Shade = False
'            PicCanvas.MouseIcon = ImageMouse(7).Picture
'            msg = "Select position for text"
            
        Case 12 '"Pen"
             ShowPenProperty = True
             
        Case 13 '"Fill"
             ShowFillProperty = True
             
        Case 15 'picture
            DrawingMode = MODE_PICTURE
            Set m_NewObject = New vbdLine
            Set new_line = m_NewObject
            new_line.IsBox = True
            new_line.m_DrawStyle = 5
            new_line.m_DrawWidth = 1
            new_line.m_FillColor = 0
            new_line.m_FillStyle = 0
            new_line.m_ForeColor = RGB(0, 0, 0)
            new_line.m_FillMode = 0
            new_line.m_TypeDraw = dPicture
            new_line.m_Shade = False
            new_line.m_Pattern = m_Pattern
            m_Pattern = ""
            'new_line.vbdObject_Bound
            MouseIcon 15
           ' msg = "Press and Hold ''Ctrl'' Button to make a Cube"
         Case 18 'pan
            OldDrawingMode = DrawingMode
            DrawingMode = MODE_PANNING
            MouseIcon 18
            Msg = "Draw mouse on the windows "
            
         Case 19 'Zoom
            OldDrawingMode = DrawingMode
            DrawingMode = MODE_START_ZOOM
            MouseIcon 19
            Msg = "Select windows for zoom"
         Case 20 '"DropperPen", "DropperFill"
           ' DeselectAll
           ' DrawingMode = MODE_NONE
           ' m_DrawingObject = False
            'm_EditObject = False
            MouseIcon 20
            Msg = "Select object"
           ' m_DrawingObject = False
          '  Redraw
    End Select
        
    ' Let the new object receive picCanvas events.
    If Not (m_NewObject Is Nothing) Then
        Set m_NewObject.canvas = PicCanvas
    End If
    
    RaiseEvent MsgControl(Msg)
    
End Sub

Private Sub MouseIcon(Key As Integer)
       
      Select Case Key
      Case 1
          PicCanvas.MouseIcon = ImageMouse(0).Picture
          'PicCanvas.MousePointer = 99
      Case 2
          PicCanvas.MouseIcon = ImageMouse(21).Picture
      Case 3
          PicCanvas.MouseIcon = ImageMouse(1).Picture
      Case 4
          PicCanvas.MouseIcon = ImageMouse(3).Picture
      Case 5
          PicCanvas.MouseIcon = ImageMouse(2).Picture
      Case 6
          PicCanvas.MouseIcon = ImageMouse(2).Picture
      Case 7
          PicCanvas.MouseIcon = ImageMouse(6).Picture
      Case 8 'pan
          PicCanvas.MouseIcon = ImageMouse(5).Picture
      Case 9
          PicCanvas.MouseIcon = ImageMouse(4).Picture
      Case 10
          PicCanvas.MouseIcon = ImageMouse(6).Picture
      Case 11
          PicCanvas.MouseIcon = ImageMouse(7).Picture
      Case 12
      Case 13
      Case 14
      Case 15
           PicCanvas.MouseIcon = ImageMouse(5).Picture
      Case 16
      Case 17
      Case 18
           PicCanvas.MouseIcon = ImageMouse(8).Picture
      Case 19
           PicCanvas.MouseIcon = ImageMouse(20).Picture
      Case 20
            PicCanvas.MouseIcon = ImageMouse(11).Picture
      Case 22 'pan2
           PicCanvas.MouseIcon = ImageMouse(22).Picture
      End Select
      'PicCanvas.MouseIcon = ImageMouse(1).Picture
      PicCanvas.MousePointer = 99
End Sub

Sub ChoiceColorForControl()
    RaiseEvent ColorSelected(1, m_FillColor)
    RaiseEvent ColorSelected(2, m_ForeColor)
End Sub

' Move this object to the front,back,forward,Backward of the scene's
' object list.
Public Function SetObjectOrder(mOrder As m_Order)
   Dim the_scene As vbdScene
        
   Set the_scene = m_TheScene
   
   Select Case mOrder
   Case BringToFront
       the_scene.MoveToFront m_SelectedObjects
   Case SendToBack
        the_scene.MoveToBack m_SelectedObjects
   Case BringFoward
       the_scene.MoveToFoward m_SelectedObjects
   Case SendBackward
       the_scene.MoveToBackward m_SelectedObjects
   End Select
   
   Set_Dirty
   
   Redraw
   
End Function

Public Sub SelectAllObject()
    Dim the_scene As vbdScene
     
    ' Save the new object.
    Set the_scene = m_TheScene
    the_scene.SelectAllObject
    
    Redraw
    Set the_scene = Nothing
End Sub

Public Sub UnSelectAllObject()
    Dim the_scene As vbdScene
     
    ' Save the new object.
    Set the_scene = m_TheScene
    the_scene.DeselectAllObject
    
    Redraw
    Set the_scene = Nothing
End Sub

Public Sub NewDraw(Optional SelectPage As Boolean = False)
    Dim cW As Single, cH As Single, cImage As String, cColor As OLE_COLOR
    
    ' Create a new, empty scene object.
    Set m_TheScene = New vbdScene

    ' No objects are selected.
    Set m_SelectedObjects = New Collection
     
    PrepareToEdit
    If SelectPage Then
        cW = CanvasWidth: cH = CanvasHeight
        cColor = BackColor
        cImage = BackImage
        If FrmCanvas.ShowForm(cW, cH, cImage, cColor) = False Then
          CanvasHeight = cH
          CanvasWidth = cW
          LockObject = False
          BackImage = cImage
          BackColor = cColor
          Unload FrmCanvas
          InitPage
          SetScaleFactor 1
          RaiseEvent ZoomChange
        End If
    End If
    Filename = ""
End Sub

' Select default values and prepare to edit.
Public Sub PrepareToEdit()
    
    ' Start at normal (pixel) scale.
    If PicCanvas.ScaleMode <> vbPixels Then
       PicCanvas.ScaleMode = vbPixels
       SetScaleFull
    End If
    ' Save the initial snapshot.
    Set m_Snapshots = New Collection
    m_CurrentSnapshot = 0
    SaveSnapshot
    
    ' Enable/disable the undo and redo menus.
    RaiseEvent EnableMenusForSelection
    
    ' Select the solid DrawStyle.
    icbDrawStyle.SelectedItem = icbDrawStyle.ComboItems(1)
    icbDrawStyle.ToolTipText = icbDrawStyle.ComboItems(1).Key
    
    ' Select the 1 pixel DrawWidth.
    icbDrawWidth.SelectedItem = icbDrawStyle.ComboItems(1)
    icbDrawWidth.ToolTipText = icbDrawStyle.ComboItems(1).Key
    'Redraw
End Sub

Public Function SaveDraw(ByVal File_name As String, ByVal file_title As String) As Boolean
Dim fnum As Integer

    On Error GoTo SaveError

    ' Open the file.
    fnum = FreeFile
    Open File_name For Output As fnum

    ' Write the scene serialization into the file.
    Print #fnum, "Page (W(" + Trim(Str(CanvasWidth)) + ") " + _
                      " H(" + Trim(Str(CanvasHeight)) + ")" + _
                      " C(" + Trim(Str(BackColor)) + "))" + vbCrLf + _
                      m_TheScene.Serialization

    ' Close the file.
    Close fnum
    
    m_FileName = File_name
    m_FileTitle = file_title

    m_DataModified = False
    SaveDraw = True
    Exit Function

SaveError:
    MsgBox "Error " & Format$(Err.Number) & " saving file " & File_name & "." & vbCrLf & Err.Description, vbCritical
    SaveDraw = False
    
End Function

'Open draw file
Public Function OpenDraw(ByVal File_name As String, ByVal file_title As String) As Boolean
    
Dim fnum As Integer
Dim txt As String
Dim token_name As String
Dim token_value As String
    Dim tmptxt As String
    
    On Error GoTo LoadError

    ' Open the file.
    fnum = FreeFile
    Open File_name For Input As fnum

    ' Read the scene serialization from the file.
    txt = Input$(LOF(fnum), fnum)

    ' Close the file.
    Close fnum
    If InStr(1, txt, "Page") > 0 Then
       'Do
        GetNamedToken txt, token_name, token_value
        If token_name = "Page" Then
          tmptxt = token_value
        Do
        GetNamedToken tmptxt, token_name, token_value
        If token_name = "W" Then
           ' CanvasWidth = CLng(token_value)
        ElseIf token_name = "H" Then
           ' CanvasHeight = CLng(token_value)
        ElseIf token_name = "C" Then
            BackColor = CLng(token_value) '
        Else
           Exit Do
         End If
       Loop
       UserControl_Resize
       End If
        txt = Replace(txt, vbCrLf, "")
    End If
    ' Initialize the scene.
    GetNamedToken txt, token_name, token_value
        
    If token_name <> "Scene" Then
        MsgBox "Error loading file " & File_name & "." & vbCrLf & "This is not a VbDraw file."
    Else
        m_TheScene.Serialization = token_value
        m_DataModified = False
    End If
   
    m_FileName = File_name
    m_FileTitle = file_title
    
   ' Save the initial snapshot.
    Set m_Snapshots = New Collection
    PrepareToEdit
    OpenDraw = True
    InitPage
    SetScaleFull
    
    Exit Function
    
LoadError:
    MsgBox "Error " & Format$(Err.Number) & " loading file " & File_name & "." & vbCrLf & Err.Description, vbCritical
    OpenDraw = False
    Exit Function
    
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Filename() As String
    Filename = m_FileName
End Property

Public Property Let Filename(ByVal New_FileName As String)
    m_FileName = New_FileName
    PropertyChanged "FileName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FileTitle() As String
    FileTitle = m_FileTitle
End Property

Public Property Let FileTitle(ByVal New_FileTitle As String)
    m_FileTitle = New_FileTitle
    PropertyChanged "FileTitle"
End Property

Public Sub DeselectAll()
       ' Deselect all objects.
       DeselectAllVbdObjects
       ClearBox
End Sub

'Fill object for draw
Private Sub DrawPen()
    Dim txt As String
    icbDrawStyle.ComboItems.Clear
    Set icbDrawStyle.ImageList = imlDrawStyles
    For i = 1 To 6
        Select Case i
        Case 1: txt = "Solid"
        Case 2: txt = "Dash"
        Case 3: txt = "Dot"
        Case 4: txt = "Dash-Dot"
        Case 5: txt = "Dash-Dot-Dot"
        Case 6: txt = "Transparent"
        End Select
        icbDrawStyle.ComboItems.Add i, txt, txt
        icbDrawStyle.ComboItems(i).Image = i
    Next i
    
    icbDrawWidth.ComboItems.Clear
    Set icbDrawWidth.ImageList = imlDrawWidths
    For i = 1 To 10
        icbDrawWidth.ComboItems.Add i, Str(i) + " point", Str(i) + " point"
        icbDrawWidth.ComboItems(i).Image = i
    Next i
End Sub


Public Property Get ShowPenProperty() As Boolean
    ShowPenProperty = m_ShowPenProperty
End Property

Public Property Let ShowPenProperty(ByVal New_ShowPenProperty As Boolean)
    m_ShowPenProperty = New_ShowPenProperty
    PropertyChanged "ShowPenProperty"
    
    If MeFormView1 = False Then
       MeForm1.Move UserControl.ScaleWidth - MeForm1.Width + 1 - VScroll1.Width, 0
      MeFormView1 = True
    End If
       
    MeForm1.Visible = m_ShowPenProperty
    
End Property

Public Property Get ShowFillProperty() As Boolean
    ShowFillProperty = m_ShowFillProperty
End Property

Public Property Let ShowFillProperty(ByVal New_ShowFillProperty As Boolean)
    m_ShowFillProperty = New_ShowFillProperty
    PropertyChanged "ShowFillProperty"
    If MeFormView2 = False Then
       MeForm2.Move UserControl.ScaleWidth - MeForm1.Width + 1 - VScroll1.Width, MeForm1.Height
       MeFormView2 = True
    End If
       
    MeForm2.Visible = m_ShowFillProperty
    
End Property

Public Property Get ShowTranformProperty() As Boolean
    ShowTranformProperty = m_ShowTranformProperty
End Property

Public Property Let ShowTranformProperty(ByVal New_ShowTranformProperty As Boolean)
    m_ShowTranformProperty = New_ShowTranformProperty
    PropertyChanged "ShowTranformProperty"
    If MeFormView3 = False Then
       MeForm3.Move UserControl.ScaleWidth - MeForm1.Width + 1 - VScroll1.Width, MeForm1.Height + MeForm2.Height
       MeFormView3 = True
    End If
    MeForm3.Visible = m_ShowTranformProperty
End Property



' Let the user scale the selected objects.
Private Sub TransformScale(X_scale As Single, Y_Scale As Single)
Dim fSelect As Boolean
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim xmid As Single
Dim ymid As Single
Dim Obj As vbdObject
Dim m(1 To 3, 1 To 3) As Single
Dim Msg As String
    
     ''Debug.Print "scale " + Str(x_scale) + " " + Str(y_scale)
     Msg = "Scale X:" + Format(X_scale, "0.0") + " Y:" + Format(Y_Scale, "0.0")
     RaiseEvent MsgControl(Msg)
 
     X_scale = X_scale / 100
     Y_Scale = Y_Scale / 100
    
    ' Bound the selected objects.
    'BoundObjects m_SelectedObjects, xmin, ymin, xmax, ymax
     xmin = XminBox
     ymin = YminBox
     xmax = XmaxBox
     ymax = YmaxBox
     
    ' Make the transformation matrix.
    Select Case m_ScaleType
    Case 1 'Left top Corner
        xmid = xmax
        ymid = ymax
    Case 2 'Middle top
        xmid = xmin
        ymid = ymax
    Case 3 'Right top Corner
        xmid = xmin
        ymid = ymax
    Case 4 'Middle Right
       xmid = xmin
       ymid = ymin
    Case 5 'Bottom Right corner
       xmid = xmin
       ymid = ymin
    Case 6 'Middle Bottom
       xmid = xmin
       ymid = ymin
    Case 7 'Left bottom corner
       xmid = xmax
       ymid = ymin
    Case 8 'Middle left
       xmid = xmax
       ymid = ymin
    Case 9
       xmid = (xmin + xmax) / 2
       ymid = (ymin + ymax) / 2
    End Select
    
    'm2ScaleAt m, X_scale, Y_Scale, xmid / m_ZoomFactor, ymid / m_ZoomFactor
    m2ScaleAt m, X_scale, Y_Scale, xmid, ymid
    
    ' Add the transformation to the selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True And Obj.ObjLock = False Then
           Obj.AddTransformation m
           Obj.MakeTransformation
           fSelect = True
        End If
    Next Obj

    ' The data has changed.
    If fSelect Then
      Set_Dirty
      Redraw
    End If
End Sub

' Let the user scale the selected objects.
Private Sub TransformSkew(X_scale As Single, Y_Scale As Single)
Dim fSelect As Boolean
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim xmid As Single
Dim ymid As Single
Dim Obj As vbdObject
Dim m(1 To 3, 1 To 3) As Single
Dim Msg As String

     ''Debug.Print "scale " + Str(x_scale) + " " + Str(y_scale)
     Msg = "Skew X:" + Format(X_scale, "0.0") + " Y:" + Format(Y_Scale, "0.0")
     RaiseEvent MsgControl(Msg)
 
     X_scale = X_scale / 100
     Y_Scale = Y_Scale / 100
    
    'Bound the selected objects.
     xmin = XminBox
     ymin = YminBox
     xmax = XmaxBox
     ymax = YmaxBox
     
    ' Make the transformation matrix.
    Select Case m_ScaleType
    Case 2 'Middle top
        xmid = xmin
        ymid = ymax
    Case 4 'Middle Right
       xmid = xmin
       ymid = ymin
    Case 6 'Middle Bottom
       xmid = xmin
       ymid = ymin
    Case 8 'Middle left
       xmid = xmax
       ymid = ymin
    Case 9 'Center
       xmid = (xmin + xmax) / 2
       ymid = (ymin + ymax) / 2
    End Select
    
    'm2SkewAt m, X_scale, Y_Scale, xmid / m_ZoomFactor, ymid / m_ZoomFactor
    m2SkewAt m, X_scale, Y_Scale, xmid, ymid
    
    ' Add the transformation to the selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True And Obj.ObjLock = False Then
           Obj.AddTransformation m
           Obj.MakeTransformation
           fSelect = True
        End If
    Next Obj
   
    ' The data has changed.
    If fSelect Then
       Set_Dirty
       Redraw
   End If
End Sub

' Let the user transform the selected objects.
Private Sub TransformPoint(X_Move As Single, Y_Move As Single)
Dim fSelect As Boolean
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim xmid As Single
Dim ymid As Single
Dim Obj As vbdObject
Dim m(1 To 3, 1 To 3) As Single
        
    ' Make the transformation matrix.
    m2Translate m, X_Move, Y_Move

    ' Add the transformation to the selected objects.
    For Each Obj In m_SelectedObjects
       If Obj.Selected = True And Obj.ObjLock = False Then
          Obj.AddTransformation m
          Obj.MakeTransformation
          fSelect = True
      End If
    Next Obj

    ' The data has changed.
    If fSelect Then
        Set_Dirty
        Redraw
    End If
End Sub

' Rotate the selected objects.
Private Sub TransformRotate(m_angle As Single, XminB As Single, YminB As Single, XmaxB As Single, YmaxB As Single)
Dim fSelect As Boolean
Dim Angle As Single
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim xmid As Single
Dim ymid As Single
Dim Obj As vbdObject
Dim m(1 To 3, 1 To 3) As Single
    
    ' Get the angle of rotation.
    Angle = m_angle * PI / 180

    ' Bound the selected objects.
    xmin = XminB
    ymin = YminB
    xmax = XmaxB
    ymax = YmaxB
    
    ' Make the transformation matrix.
    xmid = (xmin + xmax) / 2
    ymid = (ymin + ymax) / 2
    m2RotateAround m, Angle, xmid, ymid

    ' Add the transformation to the selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True And Obj.ObjLock = False Then
           Obj.Angle = Obj.Angle + m_angle
           Obj.AddTransformation m
           Obj.MakeTransformation
           fSelect = True
        End If
    Next Obj

    ' The data has changed.
    If fSelect Then
     Set_Dirty
    Redraw
    End If
End Sub

' Draw Reflect the transformed data.
Private Sub TransformMirror(rHor As Integer, rVer As Integer)
Dim fSelect As Boolean
Dim dX As Single
Dim dy As Single
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim m(1 To 3, 1 To 3) As Single
Dim i As Integer

    If Obj Is Nothing Then Exit Sub
    
    Obj.Bound xmin, ymin, xmax, ymax
    
    ' Transform the data.
     If rHor > 0 Then dX = 90 Else dX = 0
     If rVer > 0 Then dy = 90 Else dy = 0
     m2ReflectAcross m, xmin + ((xmax - xmin) / 2), ymin + ((ymax - ymin) / 2), dX, dy
    
    ' Add the transformation to the selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True And Obj.ObjLock = False Then
           Obj.AddTransformation m
           Obj.MakeTransformation
           fSelect = True
        End If
    Next Obj

    ' The data has changed.
    If fSelect Then
       Set_Dirty
       Redraw
    End If
End Sub

Private Sub mDrawRotate(LastX As Single, LastY As Single)
Dim Msg As String
Dim Ang As Single
Dim Points() As PointAPI
ReDim Points(1 To 4)

    Set Ortho = New RectAngle
    Ortho.NumPoints = 4
    Ortho.X(1) = XminBox
    Ortho.X(2) = XmaxBox
    Ortho.X(3) = XmaxBox
    Ortho.X(4) = XminBox
    Ortho.Y(1) = YminBox
    Ortho.Y(2) = YminBox
    Ortho.Y(3) = YmaxBox
    Ortho.Y(4) = YmaxBox
    Ang = m2GetAngle3P(XminBox + (XmaxBox - XminBox) / 2, YminBox + (YmaxBox - YminBox) / 2, _
                       XminBox + (XmaxBox - XminBox), m_StartY, _
                       LastX, LastY)
    Msg = "Rotate Angle:" + Format(360 - Ang, "0.0")
    RaiseEvent MsgControl(Msg)
               
    If Ang = 0 Then Exit Sub
    m2Identity m
    m2RotateAround m, (Ang * PI / 180), (XmaxBox + XminBox) / 2, (YmaxBox + YminBox) / 2
    Ortho.Transform m
    
    Points(1).X = Ortho.X(1)
    Points(1).Y = Ortho.Y(1)
    PicCanvas.PSet (Points(1).X, Points(1).Y)
    For i = 2 To 4
       Points(i).X = Ortho.X(i)
       Points(i).Y = Ortho.Y(i)
       PicCanvas.Line -(Points(i).X, Points(i).Y)
    Next
    PicCanvas.Line -(Points(1).X, Points(1).Y)
    'Ortho.Draw PicCanvas
    'Ortho.ClearTransform M
    Set Ortho = Nothing
End Sub

Public Sub ViewTransform(TypeTrans As Integer)
      If MeForm3.Visible = False Then
         MeForm3.Visible = True
         ShowTranformProperty = True
      End If
      CtrTranform1.TypeTransform = TypeTrans
End Sub


Private Sub mDrawSkew(LastX As Single, LastY As Single)
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim xmid As Single
Dim ymid As Single
Dim Msg As String
Dim Ang As Single
Dim Points() As PointAPI
ReDim Points(1 To 4)
    Set Ortho = New RectAngle
       
    Ortho.NumPoints = 4
    Ortho.X(1) = XminBox
    Ortho.X(2) = XmaxBox
    Ortho.X(3) = XmaxBox
    Ortho.X(4) = XminBox
    Ortho.Y(1) = YminBox
    Ortho.Y(2) = YminBox
    Ortho.Y(3) = YmaxBox
    Ortho.Y(4) = YmaxBox
    
     Msg = "Skew X:" + Format(LastX, "0.0") + " Y:" + Format(LastY, "0.0")
     RaiseEvent MsgControl(Msg)
 
     LastX = LastX / 100
     LastY = LastY / 100
    
    ' Bound the selected objects.
    BoundObjects m_SelectedObjects, xmin, ymin, xmax, ymax
    
    ' Make the transformation matrix.
    Select Case m_ScaleType
    Case 2 'Middle top
        xmid = xmin
        ymid = ymax
    Case 4 'Middle Right
       xmid = xmin
       ymid = ymin
    Case 6 'Middle Bottom
       xmid = xmin
       ymid = ymin
    Case 8 'Middle left
       xmid = xmax
       ymid = ymin
    Case Else
       Exit Sub
    End Select
    
    m2Identity m
    'm2SkewAt m, LastX, LastY, xmid / m_ZoomFactor, ymid / m_ZoomFactor
    m2SkewAt m, LastX, LastY, xmid, ymid
    
    Ortho.Transform m
    Points(1).X = Ortho.X(1)
    Points(1).Y = Ortho.Y(1)
    PicCanvas.PSet (Points(1).X, Points(1).Y)
    For i = 2 To 4
       Points(i).X = Ortho.X(i)
       Points(i).Y = Ortho.Y(i)
       PicCanvas.Line -(Points(i).X, Points(i).Y)
    Next
    PicCanvas.Line -(Points(1).X, Points(1).Y)
    
    Set Ortho = Nothing
End Sub

Function DelObject()
       ObjectDelete
End Function

Public Sub Set_Dirty()
      RaiseEvent SetDirty
      RaiseEvent EnableMenusForSelection
End Sub

Public Function CutObject()
    If Obj Is Nothing Then Exit Function
    
    Clipboard.Clear
    Clipboard.SetText Obj.Serialization

    'Delete object
    ObjectDelete
    
End Function

Public Function CopyObject()
     If Obj Is Nothing Then Exit Function
     Clipboard.Clear
     Clipboard.SetText Obj.Serialization
     Set_Dirty
End Function

Public Function PasteObject()
Dim NewTxt As String, OldTxt As String, token_name As String, token_value As String
    
    NewTxt = Clipboard.GetText
    OldTxt = m_TheScene.Serialization
    GetNamedToken OldTxt, token_name, token_value
    If token_name = "Scene" Then
        m_TheScene.Serialization = token_value + vbCr + NewTxt
        m_DataModified = False
    End If
    
    Set_Dirty
   
    Redraw
End Function
 
Public Sub ClearTransform()
Dim Obj As vbdObject

    For Each Obj In m_SelectedObjects
       If Obj.Selected = True Then
          Obj.Angle = 0
          Obj.ClearTransformation
       End If
    Next Obj

    ' The data has changed.
    Set_Dirty
    Redraw
End Sub

Public Function ClearObject()
       Clipboard.Clear
End Function

Public Function IsSelectObject() As Boolean
       If Obj Is Nothing Then
          IsSelectObject = False
       Else
          IsSelectObject = True
       End If
End Function

Public Sub FileExport(Filename As String)
Dim mf_dc As Long
Dim hMf As Long
Dim old_size As PointAPI

    ' Create the metafile.
    mf_dc = CreateMetaFile(ByVal Filename)
    If mf_dc = 0 Then
        MsgBox "Error creating the metafile.", vbExclamation
        Exit Sub
    End If

    ' Set the metafile's size to something reasonable.
    SetWindowExtEx mf_dc, PicCanvas.ScaleWidth, PicCanvas.ScaleHeight, old_size

    ' Draw in the metafile.
    m_TheScene.DrawInMetafile mf_dc

    ' Close the metafile.
    hMf = CloseMetaFile(mf_dc)
    If hMf = 0 Then
        MsgBox "Error closing the metafile.", vbExclamation
    End If

    ' Delete the metafile to free resources.
    If DeleteMetaFile(hMf) = 0 Then
        MsgBox "Error deleting the metafile.", vbExclamation
    End If
    If FileExists(Filename) Then
       MsgBox Filename & " Saved OK!"
    Else
       MsgBox "ERROR! " & Filename & " Not Saved!"
    End If
End Sub

Public Sub FileExportBitmap(Filename As String)
   
   ' Make picHidden big enough to hold everything.
    picHidden.Width = m_CanvasWidth 'PicCanvas.Width
    picHidden.Height = m_CanvasHeight 'PicCanvas.Height
    picHidden.ScaleLeft = 0 'PicCanvas.ScaleLeft
    picHidden.ScaleTop = 0 'PicCanvas.ScaleTop
    
    picHidden.ScaleWidth = m_CanvasWidth 'PicCanvas.ScaleWidth
    picHidden.ScaleHeight = m_CanvasHeight 'PicCanvas.ScaleHeight
    ' Erase the picture.
    picHidden.Line (picHidden.ScaleLeft, picHidden.ScaleTop)-Step(picHidden.ScaleWidth, picHidden.ScaleHeight), vbWhite, BF
    picHidden.Line (0, 0)-(m_CanvasWidth, m_CanvasHeight), RGB(0, 0, 0), B
    ' Deselect all the objects.
    DeselectAllVbdObjects
    Redraw
    picHidden.AutoRedraw = True
    ' Draw the bitmap on picHidden.
    m_TheScene.Draw picHidden, False
    picHidden.Picture = picHidden.Image
    
    ' Save the picture.
    If StartUpGDIPlus(GdiplusVersion) Then
       If SavePictureFromHDC(picHidden, Filename) Then
          MsgBox Filename & " Saved OK!", vbInformation
       Else
          MsgBox "ERROR! " & Filename & " Not Saved!", vbCritical
       End If
       ShutdownGDIPlus
    End If
        
End Sub

Public Sub PrintDraw()
    Dim OldFactor As Single
    Dim Wx1 As Long, Wx2 As Long, Wy1 As Long, Wy2 As Long
    
    If Printers.Count < 1 Then
       MsgBox "No printer", vbCritical, "Printing"
       Exit Sub
    End If
    'Read old margins
    Wx1 = Wxmin
    Wx2 = Wxmax
    Wy1 = Wymin
    Wy2 = Wymax
    'Set New margins
    Wxmin = 0  ' margins.
    Wymin = 0
    Wxmax = m_CanvasWidth
    Wymax = m_CanvasHeight
        
    mLockControl = True 'Lock control
   'OldFactor = gZoomFactor 'read zoomfactor
    ZoomFactor = 1
    picHidden.Width = m_CanvasWidth
    picHidden.Height = m_CanvasHeight
    picHidden.ScaleMode = 3
    ' Erase the picture.
    picHidden.Line (picHidden.ScaleLeft, picHidden.ScaleTop)-Step(picHidden.ScaleWidth, picHidden.ScaleHeight), vbWhite, BF

    ' Deselect all the objects.
    DeselectAllVbdObjects
    Redraw
    picHidden.AutoRedraw = True
    ' Draw on picHidden.
    m_TheScene.Draw picHidden
    picHidden.Picture = picHidden.Image
    frmPrint.ShowForm picHidden.Picture
    mLockControl = False 'Unlock Control
    
    picHidden.Picture = LoadPicture()
    picHidden.Line (picHidden.ScaleLeft, picHidden.ScaleTop)-Step(picHidden.ScaleWidth, picHidden.ScaleHeight), vbWhite, BF
    picHidden.Width = m_CanvasWidth
    picHidden.Height = m_CanvasHeight
    
    'Set old Margins
     Wxmin = Wx1
     Wxmax = Wx2
     Wymin = Wy1
     Wymax = Wy2
End Sub

Private Sub ComDropperFill_Click()
      'm_ReadFillProperty = True
      OldDrawingMode = DrawingMode
      DrawingMode = MODE_ReadFill
      SelectTool 20 '"DropperFill"
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FillColor2() As OLE_COLOR
    FillColor2 = m_FillColor2
End Property

Public Property Let FillColor2(ByVal New_FillColor2 As OLE_COLOR)
    m_FillColor2 = New_FillColor2
    ChangeFillColor 2, m_FillColor2
    PropertyChanged "FillColor2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Pattern() As String
    Pattern = m_Pattern
End Property

Public Property Let Pattern(ByVal New_Pattern As String)
    m_Pattern = New_Pattern
    ChangePattern m_Pattern
    PropertyChanged "Pattern"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Gradient() As Integer
    Gradient = m_TypeGradient
End Property

Public Property Let Gradient(ByVal New_TypeGradient As Integer)
    m_TypeGradient = New_TypeGradient
    ChangeGradient m_TypeGradient
    PropertyChanged "Gradient"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,255
Public Property Get Blend() As Integer
    Blend = m_Blend
End Property

Public Property Let Blend(ByVal New_Blend As Integer)
    m_Blend = New_Blend
    PropertyChanged "Blend"
    ChangeBlend m_Blend
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get LockObject() As Boolean
    LockObject = m_LockObject
End Property

Public Property Let LockObject(ByVal New_LockObject As Boolean)
    m_LockObject = New_LockObject
    PropertyChanged "LockObject"
    ObjectLock m_LockObject
End Property

Private Sub ObjectLock(isLock As Boolean)
      If Not Obj Is Nothing Then
        If Obj.ObjLock <> isLock Then
         Obj.ObjLock = isLock
         m_LockObject = Obj.ObjLock
         Redraw
         End If
      End If
End Sub

Private Sub ObjectDelete()
      DeletevbdObject
      ' Save the current snapshot.
      Set_Dirty
      Redraw
End Sub


'Read Path text and make PointCoolds and Type for draw
Private Sub ReadPathText(Obj As PictureBox, _
                         txt As String, _
                         Point_Coords() As PointAPI, _
                         Point_Types() As Byte, _
                         NumPoints As Long)
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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = m_hDC
End Property

Public Property Let hDC(ByVal New_hDC As Long)
    m_hDC = New_hDC
    PropertyChanged "hDC"
End Property

Sub ChangeMenu()
    If Obj Is Nothing Then
       RaiseEvent EnableMenuBitMap(False)
       RaiseEvent EnableMenuText(False)
       Exit Sub
    End If
    If Obj.TypeDraw = dPicture Then
       RaiseEvent EnableMenuBitMap(True)
    Else
       RaiseEvent EnableMenuBitMap(False)
    End If
    If Obj.TypeDraw = dText Then
        RaiseEvent EnableMenuText(True)
    Else
        RaiseEvent EnableMenuText(False)
    End If
End Sub
'
Public Property Get ObjPicture() As Picture
      If Obj Is Nothing Then Exit Property
      If Obj.TypeDraw = dPicture Then
         Set ObjPicture = Obj.Picture
      End If
End Property

Public Property Set ObjPicture(ByVal New_ObjPicture As Picture)
    If Obj.TypeDraw = dPicture Then
       Set Obj.Picture = New_ObjPicture
       PropertyChanged "ObjPicture"
       Set_Dirty
    End If
End Property

Public Sub EditText()
    Dim nfonts As New StdFont
    Dim PointCoords() As PointAPI
    Dim PointType() As Byte, tx() As Single, ty() As Single, TypePoint() As Byte
    Dim iCounter As Long, NewText As String, xmin As Single, ymin As Single, xmax As Single, ymax As Single, mAlingText As Integer

    If Obj Is Nothing Then Exit Sub
    
      If Obj.TypeDraw = dText Or Obj.TypeDraw = dTextFrame Then
          NewText = Obj.TextDraw
          nfonts.Charset = Obj.Charset
          nfonts.Italic = Obj.Italic
          nfonts.Name = Obj.Name
          nfonts.Size = Obj.Size
          nfonts.Strikethrough = Obj.Strikethrough
          nfonts.Underline = Obj.Underline
          nfonts.Weight = Obj.Weight
          If FrmFonts.ShowForm(nfonts, NewText, mAlingText) = False Then
              Obj.Bold = nfonts.Bold
              Obj.Charset = nfonts.Charset
              Obj.Italic = nfonts.Italic
              Obj.Name = nfonts.Name
              Obj.Size = nfonts.Size
              Obj.Strikethrough = nfonts.Strikethrough
              Obj.Underline = nfonts.Underline
              Obj.Weight = nfonts.Weight
              Obj.TextDraw = NewText
              With PicCanvas
                 .Font.Bold = nfonts.Bold
                 .Font.Charset = nfonts.Charset
                 .Font.Italic = nfonts.Italic
                 .Font.Name = nfonts.Name
                 .Font.Size = nfonts.Size
                 .Font.Strikethrough = nfonts.Strikethrough
                 .Font.Underline = nfonts.Underline
                 .Font.Weight = nfonts.Weight
              End With
              If Obj.TypeDraw = dText Then
                 PicCanvas.CurrentX = Obj.CurrentX
                 PicCanvas.CurrentY = Obj.CurrentY
                 ReadPathText PicCanvas, NewText, PointCoords(), PointType(), iCounter
                 ReDim tx(1 To iCounter), ty(1 To iCounter), TypePoint(1 To iCounter)
                 With m_Polygon
                    For i = 1 To iCounter
                       tx(i) = PointCoords(i - 1).X
                       ty(i) = PointCoords(i - 1).Y
                       TypePoint(i) = PointType(i - 1)
                    Next
                 End With
                 Obj.NewPoint iCounter, tx, ty, TypePoint
              End If
              UnSelectAllObject
              Redraw
          End If
      ElseIf Obj.TypeDraw = dPolygon Then
          mnueditPoints_Click
      End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get EditPoint() As Boolean
    EditPoint = m_EditPoint
End Property

Public Property Let EditPoint(ByVal New_EditPoint As Boolean)
    m_EditPoint = New_EditPoint
    SelectTool 1
    Redraw
    PropertyChanged "EditPoint"
End Property

Private Sub DrawPoint()
     Dim aa As Integer, OldFillStyle As Long, Oldcolor As Long, OldDrawStyle As Long
     Dim mp() As PointAPI, OldScale As typScaleMode
     Dim mStep As Integer
     
     mStep = IIf(GAP / gZoomFactor > 0.5, GAP / gZoomFactor, 1)
     
     If m_NumPoints = 0 Then Exit Sub
    
     If m_SelectPoint > 0 Then
       If Obj.TypeDraw = dCurve Or Obj.TypeDraw = dFreePolygon Or Obj.TypeDraw = dPolygon Or _
          Obj.TypeDraw = dPolyline Or Obj.TypeDraw = dPolydraw Or Obj.TypeDraw = dRectAngle Then
         If m_TypePoint(m_SelectPoint) = 3 Then
            m_OriginalPoints(m_SelectPoint).X = m_OriginalPoints(1).X
            m_OriginalPoints(m_SelectPoint).Y = m_OriginalPoints(1).Y
         ElseIf m_TypePoint(m_NumPoints) = 3 And Obj.TypeDraw <> dText Then
            m_OriginalPoints(m_NumPoints).X = m_OriginalPoints(1).X
            m_OriginalPoints(m_NumPoints).Y = m_OriginalPoints(1).Y
         End If
        End If
     End If
  
      OldScale = SetWordScale(PicCanvas)
      PolyDraw PicCanvas.hDC, m_OriginalPoints(1), m_TypePoint(1), m_NumPoints
      ResetWordScale PicCanvas, OldScale
      
      OldFillStyle = PicCanvas.FillStyle
      Oldcolor = PicCanvas.FillColor
      PicCanvas.FillStyle = vbFSSolid
      PicCanvas.FillColor = vbGreen
      
      aa = 0
      
      For i = 1 To m_NumPoints
          If aa = 3 Then aa = 0
          If m_TypePoint(i) = 4 Then
             aa = aa + 1
             If aa = 1 Then
                If i <= e_NumPoints Then
                   PicCanvas.Circle (m_OriginalPoints(i).X, m_OriginalPoints(i).Y), mStep, vbGreen
                   PicCanvas.Circle (m_OriginalPoints(i).X, m_OriginalPoints(i).Y), mStep
                End If
             ElseIf aa = 3 Then
                 If i <= e_NumPoints Then
                    PicCanvas.Circle (m_OriginalPoints(i - 1).X, m_OriginalPoints(i - 1).Y), mStep, vbGreen
                    PicCanvas.Circle (m_OriginalPoints(i - 1).X, m_OriginalPoints(i - 1).Y), mStep
                 End If
             ElseIf aa = 2 Then
                If i <= e_NumPoints Then
                   PicCanvas.Line (m_OriginalPoints(i + 1).X - mStep, m_OriginalPoints(i + 1).Y - mStep)-(m_OriginalPoints(i + 1).X + mStep, m_OriginalPoints(i + 1).Y + mStep), vbGreen, BF
                   PicCanvas.Line (m_OriginalPoints(i + 1).X - mStep, m_OriginalPoints(i + 1).Y - mStep)-(m_OriginalPoints(i + 1).X + mStep, m_OriginalPoints(i + 1).Y + mStep), , B
                End If
             End If
          Else
             If i <= e_NumPoints Then
                PicCanvas.Line (m_OriginalPoints(i).X - mStep, m_OriginalPoints(i).Y - mStep)-(m_OriginalPoints(i).X + mStep, m_OriginalPoints(i).Y + mStep), vbGreen, BF
                PicCanvas.Line (m_OriginalPoints(i).X - mStep, m_OriginalPoints(i).Y - mStep)-(m_OriginalPoints(i).X + mStep, m_OriginalPoints(i).Y + mStep), , B
             End If
'             End If
          End If
       Next
     
      If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
         i = m_SelectPoint
          PicCanvas.FillColor = vbRed
          PicCanvas.Line (m_OriginalPoints(i).X - mStep, m_OriginalPoints(i).Y - mStep)-(m_OriginalPoints(i).X + mStep, m_OriginalPoints(i).Y + mStep), vbRed, BF
          PicCanvas.Line (m_OriginalPoints(i).X - mStep, m_OriginalPoints(i).Y - mStep)-(m_OriginalPoints(i).X + mStep, m_OriginalPoints(i).Y + mStep), , B
      End If

      PicCanvas.DrawStyle = vbDot
      aa = 0
      List1.Clear
      For i = 1 To m_NumPoints
          If m_ShowObjectPoint Then
             List1.AddItem Trim(Str(i)) + "." + Str(m_OriginalPoints(i).X) + "-" + Str(Str(m_OriginalPoints(i).Y))
          End If
          If aa = 3 Then aa = 0
          If m_TypePoint(i) = 4 Then
             aa = aa + 1
             If aa = 1 Or aa = 3 Then
                 If i - 1 > 0 Then
                 PicCanvas.Line (m_OriginalPoints(i - 1).X, m_OriginalPoints(i - 1).Y)-(m_OriginalPoints(i).X, m_OriginalPoints(i).Y)
                 End If
             End If
          End If
       Next
              
       PicCanvas.FillStyle = OldFillStyle
       PicCanvas.FillColor = Oldcolor
       PicCanvas.DrawStyle = OldDrawStyle
End Sub

Private Sub PolyPoints(nPoint As Long, cx As Single, cy As Single)
    If nPoint > 0 Then
        m_OriginalPoints(nPoint).X = cx
        m_OriginalPoints(nPoint).Y = cy
    End If
   
End Sub

' Calculate new Point in the line.
Private Function mAddNode(ByVal X As Single, ByVal Y As Single) As PointAPI

        Dim MinX As Single, MinY As Single, i As Long, e As Long, nD As Integer
        Dim NewDist As Single, mDist As Single, aa As Long, t As Long
        Dim X1 As Single, Y1 As Single, X2 As Single, Y2 As Single
        Dim Points() As PointAPI, mTypePoint() As Byte, mp1 As PointAPI, mp2 As PointAPI, mp3 As PointAPI
        Dim mStep As Integer
        
        mStep = IIf(GAP / gZoomFactor > 0, GAP / gZoomFactor, 3)
        'add sto telos
        NewDist = 0
        e = 0
        For i = 1 To m_NumPoints - 1
            mDist = DistToSegment(X, Y, m_OriginalPoints(i).X, m_OriginalPoints(i).Y, _
                                        m_OriginalPoints(i + 1).X, m_OriginalPoints(i + 1).Y, MinX, MinY)
            If NewDist >= mDist Or NewDist = 0 Then
               NewDist = mDist
               e = i + 1
               'if On the node then find midpoint from next node
               If MinX + mStep >= m_OriginalPoints(i).X And MinX - mStep <= m_OriginalPoints(i).X Then
               If MinY + mStep >= m_OriginalPoints(i).Y And MinY - mStep <= m_OriginalPoints(i).Y Then
                  mp1 = MidPoint(m_OriginalPoints(i).X, m_OriginalPoints(i).Y, _
                           m_OriginalPoints(i + 1).X, m_OriginalPoints(i + 1).Y) ', _
                           MinX, MinY
               End If
               End If
               If m_TypePoint(e) = 4 Then
                  mp2 = MidPoint(m_OriginalPoints(i).X, m_OriginalPoints(i).Y, mp1.X, mp1.Y)  ', _
                           X1 , Y1
                  mp3 = MidPoint(mp1.X, mp1.Y, _
                           m_OriginalPoints(i + 1).X, m_OriginalPoints(i + 1).Y) ', _
                           X2, Y2
               Else
                  X1 = MinX
                  Y1 = MinY
                  X2 = MinX
                  Y2 = MinY
               End If
               mAddNode.X = mp1.X 'MinX
               mAddNode.Y = mp1.Y 'MinY
            End If
        Next

         If e >= 0 And e <= m_NumPoints Then
             'm_Canvas.Circle (AddNode.X, AddNode.Y), 5
             'Check Curver
             If m_TypePoint(e) = 4 Then nD = 3 Else nD = 1
            
             ReDim Points(1 To m_NumPoints + nD)
             ReDim mTypePoint(1 To m_NumPoints + nD)
             aa = 0
             For i = 1 To m_NumPoints + nD
                If e = i And nD = 1 Then
                  Points(i).X = mAddNode.X
                  Points(i).Y = mAddNode.Y
                  mTypePoint(i) = 2
                  
                ElseIf e = i And nD = 3 Then
                  Points(i).X = X1 'AddNode.X + 10
                  Points(i).Y = Y1 'AddNode.Y
                  mTypePoint(i) = 4
                  
                  Points(i + 1).X = mAddNode.X
                  Points(i + 1).Y = mAddNode.Y
                  mTypePoint(i + 1) = 4
                  
                  Points(i + 2).X = X2 'AddNode.X - 10
                  Points(i + 2).Y = Y2 'AddNode.Y
                  mTypePoint(i + 2) = 4
                  i = i + 2
                Else
                    aa = aa + 1
                    Points(i).X = m_OriginalPoints(aa).X
                    Points(i).Y = m_OriginalPoints(aa).Y
                    mTypePoint(i) = m_TypePoint(aa)
                End If
             Next
             m_NumPoints = m_NumPoints + nD
             m_OriginalPoints = Points
             m_TypePoint = mTypePoint
             Redim_Point
        End If

End Function

'Delete select node
Public Sub mDeleteNode()
        Dim Points() As PointAPI, aa As Long, i As Long, t As Integer
        Dim mTypePoint() As Byte, sPoint As Integer, ePoint As Integer
        Dim Arr()
        ReDim Arr(0)
        If m_NumPoints > 2 Then
           If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
                If m_SelectPoint + 1 = m_NumPoints Then
                   If m_TypePoint(m_SelectPoint + 1) = 2 Then
                      m_NumPoints = m_NumPoints - 1
                      ReDim Points(1 To m_NumPoints)
                      ReDim mTypePoint(1 To m_NumPoints)
                      For i = 1 To m_NumPoints
                         Points(i) = m_OriginalPoints(i)
                          mTypePoint(i) = m_TypePoint(i)
                      Next
                      m_OriginalPoints = Points
                      m_TypePoint = mTypePoint
                   End If
                ElseIf m_TypePoint(m_SelectPoint) = 4 Then
                    If IsControl(m_TypePoint, m_SelectPoint) = False Then
                        For i = 1 To m_NumPoints '- 1
                            For t = i To i + 3
                                If m_TypePoint(i) = 4 Then
                                    If Arr(UBound(Arr)) <> i + 2 Then
                                        ReDim Preserve Arr(UBound(Arr) + 1)
                                        Arr(UBound(Arr)) = i + 2 ' m_TypePoint(I)
                                        If Arr(UBound(Arr)) > m_NumPoints Then
                                            Arr(UBound(Arr)) = m_NumPoints
                                        End If
                                        t = t + 3
                                        i = i + 2
                                        Exit For
                                    End If
                                Else
                                    If Arr(UBound(Arr)) <> i Then
                                        ReDim Preserve Arr(UBound(Arr) + 1)
                                        Arr(UBound(Arr)) = i 'm_TypePoint(I)
                                        End If
                                    End If
                            Next
                        Next
                        For i = 1 To UBound(Arr) - 1
                            If m_SelectPoint >= Arr(i) And m_SelectPoint < Arr(i + 1) Then
                                sPoint = i
                            End If
                        Next
                        If sPoint > 0 Then ePoint = sPoint + 1 Else Exit Sub
                        
                        m_NumPoints = m_NumPoints - (Arr(ePoint) - Arr(sPoint))
                
                        ReDim Points(1 To m_NumPoints)
                        ReDim mTypePoint(1 To m_NumPoints)
                        aa = 0
                        For i = 1 To UBound(m_OriginalPoints) '+ 1
                            If i >= Arr(sPoint) And i < Arr(ePoint) Then
                            'Stop
                            Else
                                aa = aa + 1
                                Points(aa).X = m_OriginalPoints(i).X
                                Points(aa).Y = m_OriginalPoints(i).Y
                                mTypePoint(aa) = m_TypePoint(i)
                            End If
                        Next
                        If m_SelectPoint = UBound(m_TypePoint) Then mTypePoint(1) = m_TypePoint(UBound(m_TypePoint))
                        If m_SelectPoint = 1 Or mTypePoint(1) <> 6 Then mTypePoint(1) = 6
                        m_OriginalPoints = Points
                        m_TypePoint = mTypePoint
                        ''Debug.Print sPoint, sPoint + 2, ePoint
                        ' Stop
                    End If
                Else
                    m_NumPoints = m_NumPoints - 1
                    ReDim Points(1 To m_NumPoints)
                    ReDim mTypePoint(1 To m_NumPoints)
                    aa = 0
                    For i = 1 To m_NumPoints + 1
                        If m_SelectPoint <> i Then
                            aa = aa + 1
                            Points(aa).X = m_OriginalPoints(i).X
                            Points(aa).Y = m_OriginalPoints(i).Y
                            mTypePoint(aa) = m_TypePoint(i)
                        End If
                    Next
                    If m_SelectPoint = UBound(m_TypePoint) Then mTypePoint(1) = m_TypePoint(UBound(m_TypePoint))
                    If m_SelectPoint = 1 Or mTypePoint(1) <> 6 Then mTypePoint(1) = 6
                    m_OriginalPoints = Points
                    m_TypePoint = mTypePoint
                End If
                Redim_Point
           End If
           'DrawPoint
        Else
           m_NumPoints = 0
        End If
End Sub

'Make Curve to Line
Public Sub mtoLine()
        Dim Points() As PointAPI, P() As PointAPI, aa As Long, i As Long, t As Integer
        Dim mTypePoint() As Byte, sPoint As Integer, ePoint As Integer
        Dim Arr()
        ReDim Arr(0)
        If m_NumPoints > 2 Then
          If IsControl(m_TypePoint, m_SelectPoint) Then Exit Sub
          m_SelectPoint = m_SelectPoint + 1
          
          If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
             If m_SelectPoint > 1 Then m_SelectPoint = m_SelectPoint + 1
             If m_SelectPoint >= m_NumPoints Then m_SelectPoint = m_NumPoints - 1
               For i = 1 To m_NumPoints - 1
                     For t = i To i + 3
                      If m_TypePoint(i) = 4 Then
                         If Arr(UBound(Arr)) <> i + 2 Then
                         ReDim Preserve Arr(UBound(Arr) + 1)
                         Arr(UBound(Arr)) = i + 2
                         If Arr(UBound(Arr)) > m_NumPoints Then
                            Arr(UBound(Arr)) = m_NumPoints
                         End If
                         i = i + 2
                         Exit For
                         End If
                      Else
                        If Arr(UBound(Arr)) <> i Then
                        ReDim Preserve Arr(UBound(Arr) + 1)
                        Arr(UBound(Arr)) = i
                        End If
                     End If
                    Next
               Next
               For i = 1 To UBound(Arr) - 1
                   If m_SelectPoint >= Arr(i) And m_SelectPoint < Arr(i + 1) Then
                      sPoint = i
                   End If
                Next
                'If sPoint > 0 Then ePoint = sPoint + 1
                If sPoint > 0 Then ePoint = sPoint + 1 Else Exit Sub
                m_NumPoints = m_NumPoints - (((Arr(ePoint)) - (Arr(sPoint))) - 1)
                
                P = m_OriginalPoints
                For i = Arr(sPoint) + 1 To Arr(ePoint) - 1
                    P(i).X = 0
                    P(i).Y = 0
                Next
                
                ReDim Points(1 To m_NumPoints)
                ReDim mTypePoint(1 To m_NumPoints)
                aa = 0
                For i = 1 To UBound(P)
                  If P(i).X <> 0 Then
                    aa = aa + 1
                    Points(aa).X = m_OriginalPoints(i).X
                    Points(aa).Y = m_OriginalPoints(i).Y
                    mTypePoint(aa) = m_TypePoint(i)
                 End If
                Next
                mTypePoint(Arr(sPoint) + 1) = 2
                m_OriginalPoints = Points
                m_TypePoint = mTypePoint
          End If
          Redim_Point
         ' DrawPoint
        End If
End Sub

'Make Line to Curve
Public Sub mtoCurve(Optional mSelectPoint As Integer)
     Dim Points() As PointAPI, aa As Long, mType As Byte, i As Long
     Dim mTypePoint() As Byte, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single
     Dim mp1 As PointAPI, mp2 As PointAPI, mp3 As PointAPI, StarPoint As Long, EndPoint As Long
     
     If mSelectPoint = 0 And m_SelectPoint > 0 Then mSelectPoint = m_SelectPoint
     If m_NumPoints > 2 Then
       
         If mSelectPoint > 0 And mSelectPoint <= m_NumPoints Then
             ReDim Points(1 To m_NumPoints + 2)
             ReDim mTypePoint(1 To m_NumPoints + 2)
             
             
             mSelectPoint = mSelectPoint + 1
             If mSelectPoint = 1 Then mSelectPoint = 2
             mp1 = MidPoint(m_OriginalPoints(mSelectPoint).X, m_OriginalPoints(mSelectPoint).Y, _
                      m_OriginalPoints(mSelectPoint - 1).X, m_OriginalPoints(mSelectPoint - 1).Y) ', _
                      X1, Y1
             mp2 = MidPoint(m_OriginalPoints(mSelectPoint).X, m_OriginalPoints(mSelectPoint).Y, mp1.X, mp1.Y)  ', X2, Y2
             mp3 = MidPoint(mp1.X, mp1.Y, m_OriginalPoints(mSelectPoint - 1).X, m_OriginalPoints(mSelectPoint - 1).Y) ',  X3, Y3

             
'             If mSelectPoint + 1 > m_NumPoints Then
'               mSelectPoint = m_NumPoints
'               StarPoint = mSelectPoint
'               EndPoint = mSelectPoint - 1
'             Else
'             mSelectPoint = mSelectPoint + 1
'               StarPoint = mSelectPoint
'               EndPoint = mSelectPoint - 1
'             End If
'
'          '   If mSelectPoint = 1 Then mSelectPoint = 2
'             mp1 = MidPoint(m_OriginalPoints(StarPoint).X, m_OriginalPoints(StarPoint).Y, _
'                      m_OriginalPoints(EndPoint).X, m_OriginalPoints(EndPoint).Y)   ', X1 , Y1
'             mp2 = MidPoint(m_OriginalPoints(StarPoint).X, m_OriginalPoints(StarPoint).Y, mp1.X, mp1.Y) ', X2, Y2
'             mp3 = MidPoint(mp1.X, mp1.Y, m_OriginalPoints(EndPoint).X, m_OriginalPoints(EndPoint).Y)   ',  X3, Y3
             
             aa = 0
             For i = 1 To mSelectPoint + 2
                If mSelectPoint >= i - 2 And mSelectPoint <= i Then
                    aa = aa + 1
                    Points(aa).X = m_OriginalPoints(mSelectPoint).X
                    Points(aa).Y = m_OriginalPoints(mSelectPoint).Y
                    mTypePoint(aa) = 4
                Else
                    aa = aa + 1
                    Points(aa).X = m_OriginalPoints(i).X
                    Points(aa).Y = m_OriginalPoints(i).Y
                    mTypePoint(aa) = m_TypePoint(i)
                End If
             Next
             For i = mSelectPoint + 1 To m_NumPoints
                  aa = aa + 1
                  Points(aa).X = m_OriginalPoints(i).X
                  Points(aa).Y = m_OriginalPoints(i).Y
                  mTypePoint(aa) = m_TypePoint(i)
             Next
              Points(mSelectPoint).X = mp3.X 'X3
              Points(mSelectPoint).Y = mp3.Y 'Y3
              Points(mSelectPoint + 1).X = mp2.X 'X2
              Points(mSelectPoint + 1).Y = mp2.Y 'Y2
            ' Stop
             
             If m_TypePoint(m_NumPoints) = 3 Then
                ReDim Preserve Points(1 To UBound(Points))
                ReDim Preserve mTypePoint(1 To UBound(mTypePoint))
                Points(UBound(Points)) = m_OriginalPoints(1)
                mTypePoint(UBound(mTypePoint)) = 3
                'm_NumPoints = m_NumPoints + 1
             End If
             m_NumPoints = m_NumPoints + 2
             
'             If m_TypePoint(m_NumPoints) = 3 And mTypePoint(m_NumPoints) <> 3 Then
'                ReDim Preserve Points(1 To UBound(Points) + 1)
'                ReDim Preserve mTypePoint(1 To UBound(mTypePoint) + 1)
'                Points(UBound(Points)) = m_OriginalPoints(UBound(m_OriginalPoints)) '- 1
'                mTypePoint(UBound(mTypePoint)) = 3
'                'm_NumPoints = m_NumPoints + 1
'             End If
'             m_NumPoints = UBound(Points) 'm_NumPoints + 2
             m_OriginalPoints = Points
             m_TypePoint = mTypePoint
         End If
         Redim_Point
     End If
End Sub

'Close Node line
Public Sub mCloseNode()
     Dim Points() As PointAPI, aa As Long, mType As Byte, i As Long
     Dim mTypePoint() As Byte
     
     If m_NumPoints >= 2 Then
         If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
           ReDim Points(1 To m_NumPoints)
           ReDim mTypePoint(1 To m_NumPoints)
           Points = m_OriginalPoints
           mTypePoint = m_TypePoint
           For i = 2 To m_NumPoints '- 1
              If mTypePoint(i) = 6 Or mTypePoint(i) = 3 Then
                 mTypePoint(i) = 2
              End If
           Next
           m_NumPoints = m_NumPoints + 2
           'If mTypePoint(m_NumPoints) <> 3 Then
               ReDim Preserve Points(1 To m_NumPoints)
               ReDim Preserve mTypePoint(1 To m_NumPoints)
           'End If
          mTypePoint(m_NumPoints - 1) = 2
          mTypePoint(m_NumPoints) = 3
          Points(m_NumPoints) = m_OriginalPoints(1)
          Points(m_NumPoints - 1) = m_OriginalPoints(1)
          m_OriginalPoints = Points
          m_TypePoint = mTypePoint
          Redim_Point
          
         End If
     End If
End Sub

'Open Node line
Public Sub mOpenNode()
     Dim Points() As PointAPI, aa As Long, mType As Byte, i As Long
     Dim mTypePoint() As Byte
     
     If m_NumPoints >= 2 Then
         If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
           ReDim Points(1 To m_NumPoints)
           ReDim mTypePoint(1 To m_NumPoints)
           Points = m_OriginalPoints
           mTypePoint = m_TypePoint
           For i = 2 To m_NumPoints - 1
              If mTypePoint(i) = 3 Then
                 mTypePoint(i) = 2
              End If
           Next
          If mTypePoint(i) = 3 Then
             m_NumPoints = m_NumPoints - 1
             ReDim Preserve Points(1 To m_NumPoints)
             ReDim Preserve mTypePoint(1 To m_NumPoints)
          End If
          m_OriginalPoints = Points
          m_TypePoint = mTypePoint
          'DrawPoint
          Redim_Point
         End If
     End If
End Sub

'Break line in select node
Public Sub mBreakNode()
    Dim Points() As PointAPI, aa As Long, mType As Byte, i As Long
    Dim mTypePoint() As Byte, np As Integer, t As Long
    Dim Arr()
        ReDim Arr(0)
                       
       If m_NumPoints >= 2 Then
         If m_SelectPoint > 0 And m_SelectPoint <= m_NumPoints Then
           If m_SelectPoint = 1 Or m_SelectPoint = m_NumPoints Then
           
              ReDim Points(1 To m_NumPoints)
              ReDim mTypePoint(1 To m_NumPoints)
              Points = m_OriginalPoints
              mTypePoint = m_TypePoint
              If m_TypePoint(m_NumPoints) = 3 Then
                 mType = 2
                 mTypePoint(m_NumPoints) = 2
              ElseIf m_TypePoint(m_NumPoints) = 4 Then
                 mType = 2
                 mTypePoint(m_NumPoints) = 4
              Else
                 mType = 2
              End If
              m_NumPoints = m_NumPoints + 1
              ReDim Preserve Points(1 To m_NumPoints)
              ReDim Preserve mTypePoint(1 To m_NumPoints)
              Points(m_NumPoints) = m_OriginalPoints(m_SelectPoint)
              mTypePoint(m_NumPoints) = mType
           Else
              m_NumPoints = m_NumPoints + 1
              ReDim Points(1 To m_NumPoints)
              ReDim mTypePoint(1 To m_NumPoints)
               aa = 0
               aa = aa + 1
               Points(aa).X = m_OriginalPoints(1).X
               Points(aa).Y = m_OriginalPoints(1).Y
               mTypePoint(aa) = m_TypePoint(1)
'               If m_TypePoint(m_SelectPoint) = 2 Then
'                  np = 1
'               End If
                 For i = 2 To m_SelectPoint - np
              
                  aa = aa + 1
                  Points(aa).X = m_OriginalPoints(i).X
                  Points(aa).Y = m_OriginalPoints(i).Y
                  mTypePoint(aa) = m_TypePoint(i)
               Next
               aa = aa + 1
               Points(aa).X = m_OriginalPoints(m_SelectPoint).X
               Points(aa).Y = m_OriginalPoints(m_SelectPoint).Y
               mTypePoint(aa) = 6
'               aa = aa + 1
'               Points(aa).X = m_OriginalPoints(m_SelectPoint).X
'               Points(aa).Y = m_OriginalPoints(m_SelectPoint).Y
'              ' If m_TypePoint(m_SelectPoint) <> 4 Then
'                  mTypePoint(aa) = m_TypePoint(m_SelectPoint)
'              ' Else
'              '   mTypePoint(aa) = 6
'              ' End If
               For i = m_SelectPoint + 1 To m_NumPoints - 1 '- 3
                  aa = aa + 1
                  Points(aa).X = m_OriginalPoints(i).X
                  Points(aa).Y = m_OriginalPoints(i).Y
                  mTypePoint(aa) = m_TypePoint(i)
               Next
               aa = 0
               For i = 1 To m_NumPoints
                  If mTypePoint(i) = 6 Then aa = aa + 1
               Next
               If aa > 1 Then
                  If mTypePoint(m_NumPoints) = 3 Then
                     mTypePoint(m_NumPoints) = 2
                     m_NumPoints = m_NumPoints + 1
                     ReDim Preserve Points(1 To m_NumPoints)
                     ReDim Preserve mTypePoint(1 To m_NumPoints)
                     Points(m_NumPoints) = m_OriginalPoints(1)
                     mTypePoint(m_NumPoints) = 2
                   ElseIf mTypePoint(m_NumPoints) = 0 Then
                      mTypePoint(m_NumPoints) = m_TypePoint(UBound(m_TypePoint))
                  End If
               End If

            End If
             m_OriginalPoints = Points
             m_TypePoint = mTypePoint
             Redim_Point
          End If
         ' DrawPoint
        End If
End Sub

'Check select point if is Control
Private Function IsControl(ByRef lTypes() As Byte, ByVal cCount As Long) As Boolean
        Dim BezIdx As Long, id As Long
        Const PT_CLOSEFIGURE As Long = &H1
        Const PT_LINETO As Long = &H2
        Const PT_BEZIERTO As Long = &H4
        Const PT_MOVETO As Long = &H6
        If cCount = 0 Then Exit Function
        For id = 1 To cCount
            If ((lTypes(id) And PT_BEZIERTO) = 0) Then
               BezIdx = 0
            End If
            Select Case lTypes(id) And Not PT_CLOSEFIGURE
            Case PT_LINETO    ' Straight line segment
            Case PT_BEZIERTO    ' Curve segment
                  Select Case BezIdx
                  Case 0, 1   ' Bezier control handles
                      IsControl = True
                  Case 2    ' Bezier end point
                      IsControl = False
                  End Select
                  BezIdx = (BezIdx + 1) Mod 3 '//Reset counter after 3 Bezier points
            Case PT_MOVETO    ' Move current drawing point
            End Select
        Next
End Function

' Check is opening the line
Private Function IsOpening(ByRef lTypes() As Byte) As Boolean
        Dim cCount As Long, id As Long, aa As Long
        
        cCount = UBound(lTypes)
        If cCount = 0 Then Exit Function
        For id = 1 To cCount
            If lTypes(id) = 3 Then
               aa = aa + 1
            End If
        Next
        'aa = aa - 1
        If aa > 0 Then IsOpening = True Else IsOpening = False
End Function

Private Sub Redim_Point()
      Dim n As Long
      ReDim MX(1 To m_NumPoints)
      ReDim mY(1 To m_NumPoints)

      For n = 1 To m_NumPoints
          MX(n) = m_OriginalPoints(n).X
          mY(n) = m_OriginalPoints(n).Y
      Next
End Sub

' End a zoom operation early. This happens if the user starts a zoom and the selects another menu item instead of doing the zoom.
Private Sub StopZoom()
    If DrawingMode <> MODE_START_ZOOM Then Exit Sub
    DrawingMode = MODE_NONE
    PicCanvas.DrawMode = OldMode
   ' PicCanvas.MousePointer = vbDefault
End Sub

Public Sub SetScaleObject()
       If Obj Is Nothing Then Exit Sub
       
        Wxmin = XminBox - 20
        Wxmax = XmaxBox + 20
        Wymin = YminBox - 20
        Wymax = YmaxBox + 20
        
       ' Set the new world window bounds.
       SetWorldWindow
       
       Redraw
End Sub

' Change the level of magnification.
Public Sub SetScaleFactor(ByVal fact As Single)
Dim wid As Long 'Single
Dim hgt As Long 'Single
Dim mid As Long 'Single
Dim Fa As Double
     Fa = 1 / fact
   
    ' Compute the new world window size.
    wid = Fa * (Wxmax - Wxmin)
    hgt = Fa * (Wymax - Wymin)
    
    'for big
    If hgt < 20 Then
       hgt = 20
       wid = hgt / VAspect
    End If
    
    ' Center the new world window over the old.
    mid = (Wxmax + Wxmin) / 2
    Wxmin = mid - wid / 2
    Wxmax = mid + wid / 2
    
     mid = (Wymax + Wymin) / 2
    Wymin = mid - hgt / 2
    Wymax = mid + hgt / 2
    
    If CenterZoomX <> -1 Or CenterZoomY <> -1 Then
       CenterZoomX = (Wxmax + Wxmin) / 2
       CenterZoomY = (Wymax + Wymin) / 2
    End If
    ' Set the new world window bounds.
    SetWorldWindow
    Redraw
End Sub

'Adjust the world window so it is not too big, too small, off to one side, or of the wrong aspect ratio.
'Then map the world window to the viewport and force the viewport to repaint.
Private Sub SetWorldWindow()
Dim wid As Long
Dim hgt As Long
Dim xmid As Long
Dim ymid As Long
Dim Aspect As Single
Dim lpRect As RECT
Dim MX As Long

    wid = Wxmax - Wxmin
    xmid = (Wxmax + Wxmin) / 2
    hgt = Wymax - Wymin
    ymid = (Wymax + Wymin) / 2
        
    ' Make sure we're not too big or too small.
    If wid > DataMaxWid Then
        wid = DataMaxWid
    ElseIf wid < DataMinWid Then
        wid = DataMinWid
    End If
    If hgt > DataMaxHgt Then
        hgt = DataMaxHgt
    ElseIf hgt < DataMinHgt Then
        hgt = DataMinHgt
    End If

    ' Make the aspect ratio match the viewport aspect ratio.
    Aspect = hgt / wid
    If Aspect > VAspect Then
        ' Too tall and thin. Make it wider.
        wid = hgt / VAspect
    Else
        ' Too short and wide. Make it taller.
        hgt = wid * VAspect
    End If
    
    ' Compute the new coordinates
    Wxmin = xmid - wid / 2
    Wxmax = xmid + wid / 2
    Wymin = ymid - hgt / 2
    Wymax = ymid + hgt / 2
    
    ' Check that we're not off to one side.
    If wid > DataMaxWid Then
        ' We're wider than the picture. Center.
        xmid = (DataXmax + DataXmin) / 2
        Wxmin = xmid - wid / 2
        Wxmax = xmid + wid / 2
    Else
        ' Else see if we're too far to one side.
        If Wxmin < DataXmin And Wxmax < DataXmax Then
            ' Adjust to the right.
            Wxmax = Wxmax + DataXmin - Wxmin
            Wxmin = DataXmin
        End If
        If Wxmax > DataXmax And Wxmin > DataXmin Then
            ' Adjust to the left.
            Wxmin = Wxmin + DataXmax - Wxmax
            Wxmax = DataXmax
        End If
    End If
    If hgt > DataMaxHgt Then
        ' We're taller than the picture. Center.
        ymid = (DataYmax + DataYmin) / 2
        Wymin = ymid - hgt / 2
        Wymax = ymid + hgt / 2
    Else
        ' See if we're too far to top or bottom.
        If Wymin < DataYmin And Wymax < DataYmax Then
            ' Adjust downward.
            Wymax = Wymax + DataYmin - Wymin
            Wymin = DataYmin
        End If
        If Wymax > DataYmax And Wymin > DataYmin Then
            ' Adjust upward.
            Wymin = Wymin + DataYmax - Wymax
            Wymax = DataYmax
        End If
    End If
    
    ' Map the world window to the viewport.
     PicCanvas.Scale (Wxmin, Wymin)-(Wxmax, Wymax)  '0,0 on upper left
     PicCanvas.Refresh
    gZoomFactor = Round((DataYmax - DataYmin) / (Wymax - Wymin), 3)
     
    ' Reset the scroll bars.
    IgnoreSbarChange = True
    On Error Resume Next
    HScroll1.Visible = (wid < DataXmax - DataXmin)
    VScroll1.Visible = (hgt < DataYmax - DataYmin)
    If DataYmax > 2000 Then
        MX = 1
        MS = 1
    Else
        MX = 10
        MS = 9
    End If
        
    If VScroll1.Visible = True Then
        ComCorner.Visible = True
        ' The values of the scroll bars will be where the top/left of the world window should be.
        
        VScroll1.Max = MX * (DataYmax)
        VScroll1.Min = MX * (DataYmin + hgt)
        ' SmallChange moves the world window 1/10  of its width/height. Large change moves it 9/10 of its width/height.
        VScroll1.SmallChange = MX * (hgt / MX)
        VScroll1.LargeChange = MX * (MS * hgt / MX)
        ' Set the current scroll bar values.
        If MX * Wymax >= VScroll1.Min And MX * Wymax <= VScroll1.Max Then
           VScroll1.Value = MX * Wymax
        ElseIf MX * Wymax <= VScroll1.Min Then
           VScroll1.Value = VScroll1.Min
        Else
           VScroll1.Value = CenterZoomY
           CenterZoomY = -1
        End If
    End If
    If HScroll1.Visible = True Then
       ComCorner.Visible = True
       HScroll1.Min = MX * (DataXmin)
       HScroll1.Max = MX * (DataXmax - wid)
       ' SmallChange moves the world window 1/10 of its width/height. Large change moves it 9/10 of its width/height.
       HScroll1.SmallChange = MX * (wid / MX)
       HScroll1.LargeChange = MX * (MS * wid / MX)
       ' Set the current scroll bar values.
       If MX * Wxmin > HScroll1.Max Then 'Or 10 * Wxmin <= HScroll1.Min Then
          HScroll1.Value = HScroll1.Max
       ElseIf 10 * Wxmin < HScroll1.Min Then
          HScroll1.Value = HScroll1.Min
           HScroll1.Value = MX * Wxmin
       Else
           If CenterZoomX >= HScroll1.Min And CenterZoomX <= HScroll1.Max Then
          'IF CenterZoomX
           HScroll1.Value = CenterZoomX
           CenterZoomX = -1
           End If
       End If
    End If
    If VScroll1.Visible = False And HScroll1.Visible = False Then ComCorner.Visible = False
    RaiseEvent SizeCanvas(CanvasWidth, CanvasHeight)
    IgnoreSbarChange = False
    
End Sub

' Return to the default magnification scale.
Public Sub SetScaleFull(Optional RefreshScreen As Boolean = True)
    ' Reset the world window coordinates.
    Wxmin = DataXmin
    Wxmax = DataXmax
    Wymin = DataYmin
    Wymax = DataYmax
    
    ' Set the new world window bounds.
    SetWorldWindow
    If RefreshScreen Then
       Redraw
    End If
End Sub

' The vertical scroll bar has been moved. Adjust the world window.
Private Sub VScrollBarChanged()
Dim hgt As Single

    hgt = Wymax - Wymin
    Wymax = VScroll1.Value / 10
    Wymin = Wymax - hgt
    
    ' Remap the world window.
    IgnoreSbarChange = True
    SetWorldWindow
    IgnoreSbarChange = False
    Redraw
End Sub

' The horizontal scroll bar has been moved. Adjust the world window.
Private Sub HScrollBarChanged()
Dim wid As Single
    
    wid = Wxmax - Wxmin
    Wxmin = HScroll1.Value / 10
    Wxmax = Wxmin + wid
    
    ' Remap the world window.
    IgnoreSbarChange = True
    SetWorldWindow
    IgnoreSbarChange = False
    Redraw
End Sub

Private Sub InitPage()
    'margins 1.2*Height
    'margins 2*Width
    If m_CanvasWidth < m_CanvasHeight Then
       DataXmin = -1 * m_CanvasWidth
       DataYmin = -0.2 * m_CanvasHeight
       DataXmax = 2 * m_CanvasWidth
       DataYmax = 1.2 * m_CanvasHeight
    Else
       DataXmin = -0.2 * m_CanvasWidth
       DataYmin = -0.5 * m_CanvasHeight
       DataXmax = 1.2 * m_CanvasWidth
       DataYmax = 1.5 * m_CanvasHeight
    End If
    DataMinWid = 1
    DataMinHgt = 1
    DataMaxWid = (DataXmax - Abs(DataXmin)) * 10
    DataMaxHgt = (DataYmax - Abs(DataYmin)) * 10
End Sub

Private Sub GetLimitBox()
       Dim r As RECT
       If Obj Is Nothing Then Exit Sub
       GetRgnBox Obj.hRegion, r
       XminBox = r.Left
       YminBox = r.Top
       XmaxBox = r.Right
       YmaxBox = r.Bottom
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get CrossMouse() As Boolean
    CrossMouse = m_CrossMouse
End Property

Public Property Let CrossMouse(ByVal New_CrossMouse As Boolean)
    m_CrossMouse = New_CrossMouse
    LineX.Visible = m_CrossMouse
    LineY.Visible = m_CrossMouse
    PropertyChanged "CrossMouse"
    
    SaveSetting App.ProductName, "MOUSE", "CROSS", Trim(Str(m_CrossMouse))
End Property

Public Sub RedrawCross(X As Single, Y As Single)
   With LineX
      .X1 = PicCanvas.ScaleLeft
      .X2 = PicCanvas.ScaleLeft + PicCanvas.ScaleWidth
      .Y1 = Y
      .Y2 = Y
      .Visible = True
   End With
   With LineY
      .X1 = X
      .X2 = X
      .Y1 = PicCanvas.ScaleTop
      .Y2 = PicCanvas.ScaleTop + PicCanvas.ScaleHeight
      .Visible = True
   End With
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get DrawRuler() As Boolean
    DrawRuler = m_DrawRuler
End Property

Public Property Let DrawRuler(ByVal New_DrawRuler As Boolean)
    m_DrawRuler = New_DrawRuler
    PropertyChanged "DrawRuler"
    picCanvas_Paint
    SaveSetting App.ProductName, "DRAWCONTROL", "DrawRuler", Trim(Str(m_DrawRuler))
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowObjectPoint() As Boolean
    ShowObjectPoint = m_ShowObjectPoint
End Property

Public Property Let ShowObjectPoint(ByVal New_ShowObjectPoint As Boolean)
    m_ShowObjectPoint = New_ShowObjectPoint
    PropertyChanged "ShowObjectPoint"
    If MeFormView4 = False Then
       MeForm4.Move UserControl.ScaleWidth - MeForm4.Width + 1 - VScroll1.Width, 0
       
       MeFormView4 = True
    End If
    MeForm4.Visible = m_ShowObjectPoint
End Property

