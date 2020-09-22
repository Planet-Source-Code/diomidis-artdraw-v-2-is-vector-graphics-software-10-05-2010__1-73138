VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Begin VB.Form frmVbDraw 
   AutoRedraw      =   -1  'True
   Caption         =   "Art Draw "
   ClientHeight    =   8520
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10875
   Icon            =   "VbDraw.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ArtDraw.DrawControl DrawControl1 
      Height          =   5010
      Left            =   1680
      TabIndex        =   2
      Top             =   1830
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   8837
      ShowCanvasSize  =   -1  'True
   End
   Begin VB.PictureBox PicTollBar2 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   900
      Picture         =   "VbDraw.frx":5F32
      ScaleHeight     =   240
      ScaleWidth      =   3930
      TabIndex        =   8
      Top             =   1425
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.PictureBox PicTollBar3 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5430
      Picture         =   "VbDraw.frx":7D34
      ScaleHeight     =   240
      ScaleWidth      =   3930
      TabIndex        =   6
      Top             =   1110
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.PictureBox PicTollBar1 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   870
      Picture         =   "VbDraw.frx":9876
      ScaleHeight     =   240
      ScaleWidth      =   3930
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.PictureBox PicTools 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   915
      Picture         =   "VbDraw.frx":CD38
      ScaleHeight     =   240
      ScaleWidth      =   3690
      TabIndex        =   4
      Top             =   780
      Visible         =   0   'False
      Width           =   3690
   End
   Begin ArtDraw.ColorPalette ColorPalette1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   3
      Top             =   7380
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   979
   End
   Begin ArtDraw.CtrColor CtrColor1 
      Height          =   525
      Left            =   8430
      TabIndex        =   7
      Top             =   6840
      Visible         =   0   'False
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   926
      ColorFill       =   16777215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   7935
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10054
            MinWidth        =   10054
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "Mouse Position"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "Object measurement"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Picture         =   "VbDraw.frx":105BA
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12:12 ðì"
         EndProperty
      EndProperty
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
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   1429
      _CBWidth        =   10875
      _CBHeight       =   810
      _Version        =   "6.7.9816"
      Child1          =   "drawToolbar1"
      MinWidth1       =   7005
      MinHeight1      =   360
      Width1          =   6000
      NewRow1         =   0   'False
      Child2          =   "drawToolbar3"
      MinWidth2       =   3195
      MinHeight2      =   360
      Width2          =   3195
      NewRow2         =   -1  'True
      Child3          =   "drawToolbar2"
      MinWidth3       =   5595
      MinHeight3      =   360
      Width3          =   1200
      NewRow3         =   0   'False
      Begin ArtDraw.ucToolbar drawToolbar2 
         Height          =   360
         Left            =   3585
         TabIndex        =   11
         Top             =   420
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   635
      End
      Begin ArtDraw.ucToolbar drawToolbar3 
         Height          =   360
         Left            =   165
         TabIndex        =   9
         Top             =   420
         Width           =   3195
         _ExtentX        =   5609
         _ExtentY        =   767
         Begin VB.ComboBox ComboZoom 
            Height          =   315
            Left            =   1755
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "ComboZoom"
            Top             =   15
            Width           =   960
         End
      End
      Begin ArtDraw.ucToolbar drawToolbar1 
         Height          =   360
         Left            =   165
         TabIndex        =   10
         Top             =   30
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   635
      End
   End
   Begin ArtDraw.ucToolbar drawToolbar 
      Align           =   3  'Align Left
      Height          =   6570
      Left            =   0
      TabIndex        =   12
      Top             =   810
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   11589
      BarOrientation  =   1
   End
   Begin VB.Image imgPosition 
      Height          =   240
      Left            =   615
      Picture         =   "VbDraw.frx":13BA4
      Top             =   1620
      Width           =   240
   End
   Begin VB.Image imgBoxSize 
      Height          =   240
      Left            =   930
      Picture         =   "VbDraw.frx":1412E
      Top             =   1680
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileOpenSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSaveBitmapSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import ..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFileSaveBitmap 
         Caption         =   "Export ..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileSaveMetafile 
         Caption         =   "Export &Metafile..."
         Enabled         =   0   'False
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrintSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprintersetup 
         Caption         =   "Printer setup"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExitSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnusepcut 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnupaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnusepclear 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclear 
         Caption         =   "Clear ClipBoard"
      End
      Begin VB.Menu mnusepDel 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu Mnunormal 
         Caption         =   "Normal"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSimpleWireframe 
         Caption         =   "Simple Wireframe"
      End
      Begin VB.Menu mnufullscreenpreview 
         Caption         =   "-"
      End
      Begin VB.Menu mnufullscreen 
         Caption         =   "Full screen preview"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusepSymbol 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSymbol 
         Caption         =   "Symbol"
      End
      Begin VB.Menu mnusepform 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrossHairs 
         Caption         =   "Cross Hairs"
      End
      Begin VB.Menu mnuRuler 
         Caption         =   "Rulers"
      End
      Begin VB.Menu mnupenform 
         Caption         =   "Pen form"
      End
      Begin VB.Menu mnufillform 
         Caption         =   "Fill form"
      End
      Begin VB.Menu mnutransformform 
         Caption         =   "Transform form"
      End
      Begin VB.Menu mnuObjectPoint 
         Caption         =   "Object Point"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "T&ools"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuTool 
         Caption         =   "Arrrow"
         Index           =   1
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Pointer"
         Index           =   2
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Polyline"
         Index           =   3
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Polygon"
         Index           =   4
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Free Line"
         Index           =   5
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Curve"
         Index           =   6
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Rectangle"
         Index           =   7
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Polygon"
         Index           =   8
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Ellipse"
         Index           =   9
      End
   End
   Begin VB.Menu mnuText 
      Caption         =   "&Text"
      Begin VB.Menu mnuedittext 
         Caption         =   "Edit text"
      End
      Begin VB.Menu mnuexptudetext 
         Caption         =   "Extrude text"
         Enabled         =   0   'False
      End
      Begin VB.Menu sepparag 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvertparagraph 
         Caption         =   "Convert paragraph"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuConvertarttext 
         Caption         =   "Convert Art Text"
         Enabled         =   0   'False
      End
      Begin VB.Menu sepfittext 
         Caption         =   "-"
      End
      Begin VB.Menu mnufitPath 
         Caption         =   "Fit Text To Path"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnufitframe 
         Caption         =   "Fit Text to Frame"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuBitMap 
      Caption         =   "&Bitmap"
      Begin VB.Menu mnuFilter 
         Caption         =   "Filter"
         Shortcut        =   ^F
      End
      Begin VB.Menu sepbitmap 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDefinition 
         Caption         =   "Definition"
         Begin VB.Menu mnu_Definition 
            Caption         =   "Smooth"
            Index           =   1
         End
         Begin VB.Menu mnu_Definition 
            Caption         =   "Blur"
            Index           =   2
         End
         Begin VB.Menu mnu_Definition 
            Caption         =   "Sharpen"
            Index           =   3
         End
         Begin VB.Menu mnu_Definition 
            Caption         =   "Sharpen More"
            Index           =   4
         End
         Begin VB.Menu mnu_Definition 
            Caption         =   "Diffuse"
            Index           =   5
         End
         Begin VB.Menu mnu_Definition 
            Caption         =   "Diffuse More"
            Index           =   6
         End
         Begin VB.Menu mnu_Definition 
            Caption         =   "Pixelize"
            Index           =   7
         End
         Begin VB.Menu mnu_Definition 
            Caption         =   "Rect"
            Index           =   8
         End
         Begin VB.Menu mnu_Definition 
            Caption         =   "Fog"
            Index           =   9
         End
      End
      Begin VB.Menu mnuEdges 
         Caption         =   "Edges"
         Begin VB.Menu mnu_Edges 
            Caption         =   "Emboss"
            Index           =   1
         End
         Begin VB.Menu mnu_Edges 
            Caption         =   "Emboss More"
            Index           =   2
         End
         Begin VB.Menu mnu_Edges 
            Caption         =   "Engrave"
            Index           =   3
         End
         Begin VB.Menu mnu_Edges 
            Caption         =   "Engrave More"
            Index           =   4
         End
         Begin VB.Menu mnu_Edges 
            Caption         =   "Relief"
            Index           =   5
         End
         Begin VB.Menu mnu_Edges 
            Caption         =   "Edge Enhance"
            Index           =   6
         End
         Begin VB.Menu mnu_Edges 
            Caption         =   "Contour"
            Index           =   7
         End
         Begin VB.Menu mnu_Edges 
            Caption         =   "Connected Contour"
            Index           =   8
         End
         Begin VB.Menu mnu_Edges 
            Caption         =   "Neon"
            Index           =   9
         End
         Begin VB.Menu mnu_Edges 
            Caption         =   "Art"
            Index           =   10
         End
         Begin VB.Menu mnu_Edges 
            Caption         =   "Wave"
            Index           =   11
         End
         Begin VB.Menu mnu_Edges 
            Caption         =   "Crease"
            Index           =   12
         End
         Begin VB.Menu mnu_Edges 
            Caption         =   "Stranges"
            Index           =   13
         End
      End
      Begin VB.Menu mnuColors 
         Caption         =   "Colors"
         Begin VB.Menu mnu_Colors 
            Caption         =   "GreyScale"
            Index           =   1
         End
         Begin VB.Menu mnu_Colors 
            Caption         =   "Black && White"
            Index           =   2
            Begin VB.Menu mnu_Colors2 
               Caption         =   "Nearest Color"
               Index           =   1
            End
            Begin VB.Menu mnu_Colors2 
               Caption         =   "Enhanced Diffusion"
               Index           =   2
            End
            Begin VB.Menu mnu_Colors2 
               Caption         =   "Ordered Dither"
               Index           =   3
            End
            Begin VB.Menu mnu_Colors2 
               Caption         =   "Floyd-Steinberg"
               Index           =   4
            End
            Begin VB.Menu mnu_Colors2 
               Caption         =   "Burke"
               Index           =   5
            End
            Begin VB.Menu mnu_Colors2 
               Caption         =   "Stucki"
               Index           =   6
            End
         End
         Begin VB.Menu mnu_Colors 
            Caption         =   "Negative"
            Index           =   3
         End
         Begin VB.Menu mnu_Colors 
            Caption         =   "Swap Colors"
            Index           =   4
            Begin VB.Menu mnu_Colors4 
               Caption         =   "RGB -> BRG"
               Index           =   1
            End
            Begin VB.Menu mnu_Colors4 
               Caption         =   "RGB -> GBR"
               Index           =   2
            End
            Begin VB.Menu mnu_Colors4 
               Caption         =   "RGB -> RBG"
               Index           =   3
            End
            Begin VB.Menu mnu_Colors4 
               Caption         =   "RGB -> BGR"
               Index           =   4
            End
            Begin VB.Menu mnu_Colors4 
               Caption         =   "RGB -> GRB"
               Index           =   5
            End
         End
         Begin VB.Menu mnu_Colors 
            Caption         =   "Aqua"
            Index           =   5
         End
         Begin VB.Menu mnu_Colors 
            Caption         =   "Add Noise"
            Index           =   6
         End
         Begin VB.Menu mnu_Colors 
            Caption         =   "Gamma Correction"
            Index           =   7
         End
         Begin VB.Menu mnu_Colors 
            Caption         =   "Sepia"
            Index           =   8
         End
         Begin VB.Menu mnu_Colors 
            Caption         =   "Ice"
            Index           =   9
         End
         Begin VB.Menu mnu_Colors 
            Caption         =   "Comic"
            Index           =   10
         End
      End
      Begin VB.Menu mnuIntensity 
         Caption         =   "Intensity"
         Begin VB.Menu mnu_Intensity 
            Caption         =   "Brighter"
            Index           =   1
            Shortcut        =   +^{F3}
         End
         Begin VB.Menu mnu_Intensity 
            Caption         =   "Darker"
            Index           =   2
            Shortcut        =   +^{F4}
         End
         Begin VB.Menu mnu_Intensity 
            Caption         =   "Increase Contrast"
            Index           =   3
            Shortcut        =   +^{F5}
         End
         Begin VB.Menu mnu_Intensity 
            Caption         =   "Decrease Contrast"
            Index           =   4
            Shortcut        =   +^{F6}
         End
         Begin VB.Menu mnu_Intensity 
            Caption         =   "Dilate"
            Index           =   5
         End
         Begin VB.Menu mnu_Intensity 
            Caption         =   "Erode"
            Index           =   6
         End
         Begin VB.Menu mnu_Intensity 
            Caption         =   "Contrast Stretch"
            Index           =   7
         End
         Begin VB.Menu mnu_Intensity 
            Caption         =   "Increase Saturation"
            Index           =   8
            Shortcut        =   +^{F11}
         End
         Begin VB.Menu mnu_Intensity 
            Caption         =   "Decrease Saturation"
            Index           =   9
            Shortcut        =   +^{F12}
         End
      End
      Begin VB.Menu mnuOther 
         Caption         =   "Effect"
         Begin VB.Menu mnu_Other 
            Caption         =   "Grid 3d "
            Index           =   1
         End
         Begin VB.Menu mnu_Other 
            Caption         =   "Mirror Right to Left "
            Index           =   2
         End
         Begin VB.Menu mnu_Other 
            Caption         =   "Mirror Left to Right"
            Index           =   3
         End
         Begin VB.Menu mnu_Other 
            Caption         =   "Mirror Down to Top"
            Index           =   4
         End
         Begin VB.Menu mnu_Other 
            Caption         =   "Mirror Top to Down"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnuArrange 
      Caption         =   "&Arrange"
      Begin VB.Menu mnutransform 
         Caption         =   "&Transform"
         Begin VB.Menu mnuMove 
            Caption         =   "&Position"
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu mnuTransformRotate 
            Caption         =   "&Rotate"
            Shortcut        =   ^{F6}
         End
         Begin VB.Menu mnuTransformScale 
            Caption         =   "&Scale"
            Shortcut        =   ^{F7}
         End
         Begin VB.Menu mnuskew 
            Caption         =   "S&kew"
            Shortcut        =   ^{F8}
         End
         Begin VB.Menu mnuReflect 
            Caption         =   "&Mirror"
            Shortcut        =   ^{F9}
         End
      End
      Begin VB.Menu mnuTransformClear 
         Caption         =   "&Clear Transformations"
      End
      Begin VB.Menu mnusepSelectAll 
         Caption         =   "-"
      End
      Begin VB.Menu mnuselectall 
         Caption         =   "Select All"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuunselectall 
         Caption         =   "UnSelect all"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuSepBringFront 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArrangeSendToFront 
         Caption         =   "&Bring To Front"
         Enabled         =   0   'False
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuArrangeSendToForward 
         Caption         =   "&Bring To Forward"
         Enabled         =   0   'False
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuArrangeSendToBackward 
         Caption         =   "&Send To Backward"
         Enabled         =   0   'False
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuArrangeSendToBack 
         Caption         =   "&Send To Back"
         Enabled         =   0   'False
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnusepLock 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLock 
         Caption         =   "Lock Object"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuunloackobject 
         Caption         =   "Unlock Object"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUnlockAllObject 
         Caption         =   "Unlock All Object"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnudownload 
         Caption         =   "Download"
         Begin VB.Menu mnuDownLoad1 
            Caption         =   "Texture 1"
            Index           =   1
         End
         Begin VB.Menu mnuDownLoad1 
            Caption         =   "Texture 2"
            Index           =   2
         End
         Begin VB.Menu mnuDownLoad1 
            Caption         =   "Texture 3"
            Index           =   3
         End
         Begin VB.Menu mnuDownLoad1 
            Caption         =   "Fonts-Symbols 1"
            Index           =   4
         End
         Begin VB.Menu mnuDownLoad1 
            Caption         =   "Fonts-Symbols 2"
            Index           =   5
         End
         Begin VB.Menu mnuDownLoad1 
            Caption         =   "Fonts-Symbols 3"
            Index           =   6
         End
      End
      Begin VB.Menu sephelp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmVbDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(c) 2007 Diomidisk

Option Explicit
' The currently selected colors.
Private m_ForeColor As Integer
Private m_BackColor As Integer

' MRU list file names.
Private m_MruList As Collection
'

' Return True if it is safe to discard the current picture.
Private Function DataSafe() As Boolean
    If Not m_DataModified Then
        DataSafe = True
    Else
        Select Case MsgBox("The data has been modified. Do you want to save the changes?", vbYesNoCancel + vbInformation)
            Case vbYes
                mnuFileSave_Click
                DataSafe = Not m_DataModified
            Case vbNo
                DataSafe = True
            Case vbCancel
                DataSafe = False
        End Select
    End If
End Function

' Add this file name to the MRU list.
Private Sub MruAddName(ByVal File_name As String)
Dim i As Integer

    ' Remove any duplicates.
    For i = m_MruList.Count To 1 Step -1
        If m_MruList(i) = File_name Then
            m_MruList.Remove i
        End If
    Next i

    ' Add the new name at the front.
    If m_MruList.Count = 0 Then
        m_MruList.Add File_name
    Else
        m_MruList.Add File_name, , 1
    End If

    ' Only keep 4.
    Do While m_MruList.Count > 4
        m_MruList.Remove 5
    Loop

    ' Save the MRU list in the registry.
    For i = 1 To m_MruList.Count
        SaveSetting App.Title, "MRU", Format$(i), m_MruList(i)
    Next i
    For i = m_MruList.Count + 1 To 4
        SaveSetting App.Title, "MRU", Format$(i), ""
    Next i

    ' Display the MRU list.
    MruDisplay
End Sub
' Display the MRU list.
Private Sub MruDisplay()
Dim i As Integer

    mnuFileMRU(0).Visible = (m_MruList.Count > 0)
    For i = 1 To m_MruList.Count
        If i > mnuFileMRU.UBound Then
            Load mnuFileMRU(i)
        End If
        mnuFileMRU(i).Caption = "&" & _
            Format$(i) & " " & m_MruList(i)
        mnuFileMRU(i).Visible = True
    Next i
End Sub
' Load the MRU list.
Private Sub MruLoad()
Dim i As Integer
Dim File_name As String

    Set m_MruList = New Collection
    For i = 1 To 4
        File_name = GetSetting(App.Title, "MRU", _
            Format$(i), "")
        If Len(File_name) > 0 Then
            m_MruList.Add File_name
        End If
    Next i

    ' Display the list.
    MruDisplay
End Sub

' Flag the data as modified.
Private Sub SetDirty()
    If Not m_DataModified Then
        Caption = App.Title & "*[" & DrawControl1.FileTitle & "]"
    End If

    ' Save the current snapshot.
    SaveSnapshot
     
    m_DataModified = True
End Sub

' Set the file's name.
Private Sub SetFileName(ByVal File_name As String, ByVal file_title As String)
    ' Save the file's name and title.
    DrawControl1.Filename = File_name
    DrawControl1.FileTitle = file_title
    mnuFileSave.Enabled = Len(DrawControl1.FileTitle) > 0

    ' Update the caption.
    Caption = App.Title & " [" & DrawControl1.FileTitle & "]"

    ' Add the name to the MRU list.
    If Len(DrawControl1.Filename) > 0 Then MruAddName DrawControl1.Filename
    
End Sub

' Cancel adding an object to the collection.
Public Sub CancelObject()
    Set m_NewObject = Nothing

    ' Select the arrow tool.
    SelectArrowTool
End Sub

' Restore the previous snapshot.
Private Sub Undo()
Dim token_name As String
Dim token_value As String

    If m_CurrentSnapshot <= 1 Then Exit Sub
    Screen.MousePointer = 11
    ' Restore the previous snapshot.
    m_CurrentSnapshot = m_CurrentSnapshot - 1
    GetNamedToken m_Snapshots(m_CurrentSnapshot), token_name, token_value
    m_TheScene.Serialization = token_value
    
    ' Enable/disable the undo and redo menus.
    DrawControl1_EnableMenusForSelection 'SetUndoMenus
    DrawControl1.Redraw
    Screen.MousePointer = 0
End Sub

' Reapply a previously undone snapshot.
Private Sub Redo()
Dim token_name As String
Dim token_value As String
    
    If m_Snapshots Is Nothing Then Exit Sub
     
    If m_CurrentSnapshot >= m_Snapshots.Count Then Exit Sub
    Screen.MousePointer = 11
    ' Restore the previous snapshot.
    m_CurrentSnapshot = m_CurrentSnapshot + 1

    GetNamedToken m_Snapshots(m_CurrentSnapshot), token_name, token_value
    m_TheScene.Serialization = token_value
    
    ' Enable/disable the undo and redo menus.
    DrawControl1_EnableMenusForSelection 'SetUndoMenus
    DrawControl1.Redraw
    Screen.MousePointer = 0
End Sub

Private Sub ColorPalette1_ColorOver(cColor As Long)
    Dim stmp As String
    stmp = Right("000000" & Hex(cColor), 6)
    stmp = "R:" + Str(Int("&H" & Right$(stmp, 2))) + " - G:" + Str(Int("&H" & mid$(stmp, 3, 2))) + " - B:" + Str(Int("&H" & Left$(stmp, 2)))
    StatusBar1.Panels(3).Text = stmp
End Sub

Private Sub ColorPalette1_ColorSelected(Button As Integer, cColor As Long)
    If cColor <> -1 Then
       Select Case Button
       Case 1
           CtrColor1.ColorFill = cColor
           DrawControl1.FillStyle = 0
       Case 2
           CtrColor1.ColorBorder = cColor
           DrawControl1.DrawStyle = 0
       End Select
      
       CtrColor1.Redraw
       DrawControl1.ForeColor = CtrColor1.ColorBorder
       DrawControl1.FillColor = CtrColor1.ColorFill
       
      ' StatusBar1.Panels(4).Picture = LoadPicture()
       StatusBar1.Panels(4).Picture = CtrColor1.Image
    Else
       Select Case Button
       Case 1
           DrawControl1.FillStyle = 1
       Case 2
           DrawControl1.DrawStyle = 5
       End Select
    End If
    DrawControl1.Redraw
    
End Sub

Private Sub ComboZoom_Click()
    Dim mZoom As Single
    On Error Resume Next
    If ComboZoom.ListIndex = -1 Then Exit Sub
    mZoom = Val(ComboZoom.Text) / 100
    If mZoom > 4 Then mZoom = 4
    If mZoom < 0.1 Then mZoom = 0.1
     DrawControl1.SetScaleFull
     DrawControl1.SetScaleFactor mZoom
     StatusBar1.Panels(1).Text = "Zoom (" & Round(mZoom * 100) & "%)"
     DrawControl1.SetFocus
End Sub

Private Sub ComboZoom_KeyUp(KeyCode As Integer, Shift As Integer)
     Dim mZoom As Single
     If KeyCode = 13 Then
        mZoom = Val(ComboZoom.Text) / 100
        If mZoom > 4 Then mZoom = 4
        If mZoom < 0.1 Then mZoom = 0.1
        
        ComboZoom.Text = Trim(Str(Round(mZoom * 100))) & " %"
        StatusBar1.Panels(1).Text = "Zoom (" & Round(mZoom * 100) & "%)"
        DrawControl1.SetScaleFull False
        DrawControl1.SetScaleFactor mZoom
        DrawControl1.SetFocus
     End If
End Sub

Private Sub DrawControl1_ColorSelected(ByVal tColor As Integer, ByVal cColor As Long)
    Select Case tColor
    Case 1
       CtrColor1.ColorFill = cColor
    Case 2
       CtrColor1.ColorBorder = cColor
    End Select
    CtrColor1.Redraw
   ' StatusBar1.Panels(4).Picture = LoadPicture()
    StatusBar1.Panels(4).Picture = CtrColor1.Image
End Sub

Private Sub DrawControl1_EnableMenuBitMap(ByVal MenuOn As Boolean)
       mnuBitMap.Enabled = MenuOn
End Sub

' Enable the appropriate transformation menus.
Public Sub DrawControl1_EnableMenusForSelection()

Dim objects_selected As Boolean

    objects_selected = (m_SelectedObjects.Count > 0)
    
    mnuArrangeSendToFront.Enabled = objects_selected
    mnuArrangeSendToBack.Enabled = objects_selected
    mnuArrangeSendToForward.Enabled = objects_selected
    mnuArrangeSendToBackward.Enabled = objects_selected
    
    mnuTransformClear.Enabled = objects_selected
    mnuTransformRotate.Enabled = objects_selected
    mnuTransformScale.Enabled = objects_selected
    mnuskew.Enabled = objects_selected
    mnuReflect.Enabled = objects_selected
    mnuMove.Enabled = objects_selected
    
    drawToolbar2.EnableButton 3, objects_selected
    drawToolbar2.EnableButton 4, objects_selected
    drawToolbar2.EnableButton 5, objects_selected
    drawToolbar2.EnableButton 6, objects_selected
    drawToolbar2.CheckButton 7, DrawControl1.LockObject
    
    mnuEditUndo.Enabled = (m_CurrentSnapshot > 1)
    mnuEditRedo.Enabled = (m_CurrentSnapshot < m_Snapshots.Count)
     
    drawToolbar1.EnableButton 9, mnuEditUndo.Enabled
    drawToolbar1.EnableButton 10, mnuEditRedo.Enabled
    
    mnuEdit_Click
    
End Sub

Private Sub DrawControl1_EnableMenuText(ByVal MenuOn As Boolean)
       mnuText.Enabled = MenuOn
End Sub

Private Sub DrawControl1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
      Dim txtSM As String
      Dim tx As String
      Dim ty As String
      Dim tW As String
      Dim tH As String
           
      txtSM = " " + ScalePage(gScaleMode)
      X = Round(ScaleX(X, vbPixels, gScaleMode), 2)
      Y = Round(ScaleY(Y, vbPixels, gScaleMode), 2)
      StatusBar1.Panels(2).Text = "X:" + Format(X, "0.00") + " - Y:" + Format(Y, "0.00") + txtSM
     
      If DrawControl1.XmaxBox - DrawControl1.XminBox <> 0 Or DrawControl1.YmaxBox - DrawControl1.YminBox <> 0 Then
         tx = Str(Round(ScaleX(DrawControl1.XminBox, vbPixels, gScaleMode), 2))
         ty = Str(Round(ScaleX(DrawControl1.YminBox, vbPixels, gScaleMode), 2))
         tW = Str(Round(ScaleX(DrawControl1.XmaxBox - DrawControl1.XminBox, vbPixels, gScaleMode), 2))
         tH = Str(Round(ScaleX(DrawControl1.YmaxBox - DrawControl1.YminBox, vbPixels, gScaleMode), 2))
         StatusBar1.Panels(3).Text = "X:" + tx + " - Y:" + ty + " - W:" + tW + " - H:" + tH + txtSM
      Else
         StatusBar1.Panels(3).Text = ""
      End If
      
End Sub

Private Sub DrawControl1_MsgControl(ByVal txt As String)
     StatusBar1.Panels(1).Text = txt
End Sub

Public Sub DrawControl1_SetDirty()
      SetDirty
      m_DataModified = True
End Sub

Private Sub DrawControl1_SizeCanvas(ByVal Width As Single, ByVal Height As Single)
      Dim sW As Long, sH As Long
      If gScaleMode = 0 Then gScaleMode = 3
      sW = Round(ScaleX(Width, vbPixels, gScaleMode), 0)
      sH = Round(ScaleY(Height, vbPixels, gScaleMode), 0)
     StatusBar1.Panels(5).Text = Trim(Str(sW)) + " x " + Trim(Str(sH)) + " " + ScalePage(gScaleMode)
     StatusBar1.Panels(5).ToolTipText = "Page"
End Sub

Private Sub DrawControl1_ZoomChange()
     ComboZoom.Text = Trim(Str(Round(gZoomFactor * 100))) & " %"
     StatusBar1.Panels(1).Text = "Zoom (" & Round(gZoomFactor * 100) & "%)"
End Sub


Private Sub drawToolbar_ButtonClick(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xLeft As Long, ByVal yTop As Long)
   
    DrawControl1.SelectTool Index
    
    If drawToolbar.GetTooltip(Index) = "Pen" Or drawToolbar.GetTooltip(Index) = "Fill" Then 'Or Index = 2 Then
         ' Select the arrow tool.
          SelectArrowTool
    End If
End Sub

Private Sub drawToolbar1_ButtonClick(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xLeft As Long, ByVal yTop As Long)
    Select Case Index
    Case 1
       mnuFileNew_Click
    Case 2
       mnuFileOpen_Click
    Case 3
       mnuFileSave_Click
    Case 4
       mnuFileSaveBitmap_Click
    Case 5
       mnuPrint_Click
    Case 6
       mnuCut_Click
    Case 7
      MnuCopy_Click
    Case 8 '
       mnuPaste_Click
    Case 9
       Undo
    Case 10
       Redo
    Case 11
       mnuDelete_Click
    Case 12
       mnuSymbol_Click
    Case 13
      mnupenform_Click
    Case 14
      mnufillform_Click
    Case 15
      mnutransformform_Click
    Case 16
      mnuObjectPoint_Click
    End Select
End Sub

Private Sub drawToolbar2_ButtonClick(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xLeft As Long, ByVal yTop As Long)
    Select Case Index
    Case 1
        DrawControl1.SelectAllObject
    Case 2
       DrawControl1.UnSelectAllObject
    Case 3
        DrawControl1.SetObjectOrder BringToFront
    Case 4
        DrawControl1.SetObjectOrder SendToBack
    Case 5
        DrawControl1.SetObjectOrder BringFoward
    Case 6
        DrawControl1.SetObjectOrder SendBackward
    Case 7
        DrawControl1.LockObject = Not DrawControl1.LockObject
    Case 8
        'DrawControl1.GroupObjects
    Case 9
        'DrawControl1.UnGroupObjects
    Case 10
        'DrawControl1.AlignSelectedObjects mLeft
    Case 11
        'DrawControl1.AlignSelectedObjects mCenterV
    Case 12
        'DrawControl1.AlignSelectedObjects mRight
    Case 13
        'DrawControl1.AlignSelectedObjects mTop
    Case 14
        'DrawControl1.AlignSelectedObjects mCenterH
    Case 15
        'DrawControl1.AlignSelectedObjects mBottom
    Case 16
        'DrawControl1.AlignSelectedObjects mCenterVH
    End Select
   
End Sub

Private Sub drawToolbar3_ButtonClick(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xLeft As Long, ByVal yTop As Long)
    'If gZoomLock = True Then Exit Sub
    Select Case Index
    Case 1 'Windows
        DrawControl1.SelectTool 19
    Case 2 'Full
        DrawControl1.SetScaleFull
    Case 3 '-
        DrawControl1.SetScaleFactor 0.9
    Case 4 '+
        DrawControl1.SetScaleFactor 1.1
    Case 5 'object
        DrawControl1.SetScaleObject
    Case 9 'Pan
      DrawControl1.SelectTool 18
    End Select
    If Index <> 6 Then
       ComboZoom.Text = Trim(Str(Round(gZoomFactor * 100)) & " %")
       StatusBar1.Panels(1).Text = "Zoom (" & Round(gZoomFactor * 100) & "%)"
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
      Dim ShiftDown, AltDown, CtrlDown
      ShiftDown = (Shift And vbShiftMask) > 0
      AltDown = (Shift And vbAltMask) > 0
      CtrlDown = (Shift And vbCtrlMask) > 0

'      If m_EditObject Is Nothing Then
      Dim Index As Integer
      Select Case KeyCode
      Case vbKeyAdd
         drawToolbar3_ButtonClick 4, 0, 0, 0
      Case vbKeySubtract
         
         drawToolbar3_ButtonClick 3, 0, 0, 0
      Case vbKeyF1
           If CtrlDown Then
              drawToolbar3_ButtonClick 5, 0, 0, 0
           End If
      End Select
End Sub

Private Sub Form_Load()
    Dim Filename As String, FileTitle As String

    InitScreen
         
    If Command$ <> "" Then
       Filename = Command$
       If FileExists(Filename) Then
          SplitPath Filename, , , FileTitle
          OpenDraw Filename, FileTitle
       End If
    End If
    Me.Show
'    If App.LogMode <> 1 Then
'       MsgBox "Compile me for more speed!", vbInformation
'    End If
    Wait 1000
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = (Not DataSafe())
    If m_FormSymbolView = True Then Unload FrmSymbols
    If Cancel = 0 Then Clipboard.Clear

End Sub

Private Sub Form_Resize()
Dim wid As Single
Dim hgt As Single
  
    wid = ScaleWidth - drawToolbar.Width 'CoolBar2.Width
    If wid < 3000 Then wid = 3000
    hgt = ScaleHeight - StatusBar1.Height - CoolBar1.Height
    If hgt < 3000 Then hgt = 3000
    
    DrawControl1.Move drawToolbar.Width, CoolBar1.Height, wid, hgt - ColorPalette1.Height
    If Me.WindowState = 1 Then
       If m_FormSymbolView = True Then
          FrmSymbols.Hide
       End If
    Else
       If m_FormSymbolView = True Then
          FrmSymbols.Show
       End If
    End If
    
End Sub


Private Sub mnu_Colors_Click(Index As Integer)
     
     Screen.MousePointer = 11
     Select Case Index
     Case 1 '  Grey Scale = 12              '------
           EditObjFilter2 iGreyScale
     Case 3 '  Negative = 2                 '------
          EditObjFilter2 iNegative
     Case 5 'Aqua = 24                    '-----
          EditObjFilter2 iAqua
     Case 6 ' Add Noise = 29               '0
          EditObjFilter2 iAddNoise, 50
     Case 7 ' Gamma Correction = 31        '1-100
          EditObjFilter2 iGamma, 40
     Case 8 ' sepia =44
          EditObjFilter2 iSepia, 0
     Case 9 'Ice
          EditObjFilter2 iIce, 0
     Case 10
          EditObjFilter2 iComic, 0
     Case Else
     
         Exit Sub
     End Select
     
     DrawControl1.Redraw
     Screen.MousePointer = 0
End Sub

Private Sub mnu_Colors2_Click(Index As Integer)
     Dim nFilter As Integer
     Screen.MousePointer = 11
     ' Black & White
     Select Case Index
     Case 1 ' Nearest Color = 18       '--RGB--
          nFilter = 18
          EditObjFilter2 nFilter, RGB(180, 180, 180)
     Case 2 ' Enhanced Diffusion = 19  '-----
          nFilter = 19
          EditObjFilter2 nFilter, 0
     Case 3 'Ordered Dither = 20      '-----
          nFilter = 20
          EditObjFilter2 nFilter, 0
     Case 4 ' Floyd -Steinberg = 21    '1-n
          nFilter = 21
          EditObjFilter2 nFilter, 15
     Case 5 ' Burke = 22               '1-n
          nFilter = 22
          EditObjFilter2 nFilter, 15
     Case 6 ' Stucki = 23              '1-n
          nFilter = 23
          EditObjFilter2 nFilter, 15
     Case Else
        Exit Sub
     End Select
     
     DrawControl1.Redraw
     Screen.MousePointer = 0
End Sub

Private Sub mnu_Colors4_Click(Index As Integer)
     Dim nFilter As Integer
     Screen.MousePointer = 11
     'Swap Colors
     Select Case Index
     Case 1 ' RGB -> BRG= 16
          nFilter = 16
          EditObjFilter2 nFilter, 1
     Case 2 ' RGB -> GBR= 16
         nFilter = 16
         EditObjFilter2 nFilter, 2
     Case 3 ' RGB -> RBG= 16
         nFilter = 16
         EditObjFilter2 nFilter, 3
     Case 4 ' RGB -> BGR= 16
         nFilter = 16
         EditObjFilter2 nFilter, 4
     Case 5 ' RGB -> GRB= 16
         nFilter = 16
         EditObjFilter2 nFilter, 5
     Case Else
        Exit Sub
     End Select
     DrawControl1.Redraw
     Screen.MousePointer = 0
End Sub

Private Sub mnu_Definition_Click(Index As Integer)
     Dim nFilter As Integer, Pix As Long
     
     Screen.MousePointer = 11
     Select Case Index
     Case 1 ' Smooth = 5                    '------
          EditObjFilter2 iSmooth
     Case 2 ' Blur = 3                      '------
          EditObjFilter2 iBlur
     Case 3 ' Sharpen = 1                 '2  '0-N +-
          EditObjFilter2 iSharpen, 2
     Case 4 ' Sharpen More = 1            '0  '0-N +-
          EditObjFilter2 iSharpen, 0
     Case 5 ' Diffuse = 4                   '6
          EditObjFilter2 iDiffuse, 6
     Case 6 ' Diffuse More                 '12
          EditObjFilter2 iDiffuse, 12
     Case 7 ' Pixelize = 15                 '--size Pix
          Pix = Val(InputBox("Number pixel", "Pixelize", 5))
          If Pix > 0 Then EditObjFilter2 iPixelize, Pix
     Case 8
          EditObjFilter2 iRects, 0
     Case 9
          EditObjFilter2 iFog, 0
     Case Else
        Exit Sub
     End Select
     DrawControl1.Redraw
     Screen.MousePointer = 0
End Sub

Private Sub mnu_Edges_Click(Index As Integer)
     Dim nColor As Long
     Screen.MousePointer = 11
     
     nColor = RGB(180, 180, 180)
     
     Select Case Index
     Case 1 ' Emboss = 8                    '--RGB----
           If ShowColor(nColor) = True Then
              EditObjFilter2 iEmboss, nColor
           End If
     Case 2 ' Emboss More = 9               '--RGB----
           If ShowColor(nColor) = True Then
           EditObjFilter2 iEmbossMore, nColor
           End If
     Case 3 ' Engrave = 10                  '--RGB----
           If ShowColor(nColor) = True Then
           EditObjFilter2 iEngrave, nColor
           End If
     Case 4 ' Engrave More = 11             '--RGB----
           If ShowColor(nColor) = True Then
              EditObjFilter2 iEngraveMore, nColor
           End If
     Case 5 ' Relief = 13                   '---------
           EditObjFilter2 iRelief, 0
     Case 6 ' edge Enhance = 6              '0-2 +-
           EditObjFilter2 iEDGE, 2
     Case 7 ' Contour = 7                   '--RGB----
           If ShowColor(nColor) = True Then
              EditObjFilter2 iContour, nColor
           End If
     Case 8 ' Connected Contour = 27        '-------
           EditObjFilter2 iConnection, 0
     Case 9 ' Neon = 32                     '-------
           EditObjFilter2 iNeon, 0
     Case 10
           EditObjFilter2 iArt, 0
     Case 11
           EditObjFilter2 iSnow, 50
     Case 12
            EditObjFilter2 iWave, 5
     Case 13
            EditObjFilter2 iCrease, 512
     Case Else
          Exit Sub
     End Select
     DrawControl1.Redraw
     Screen.MousePointer = 0
End Sub

Private Sub mnu_Intensity_Click(Index As Integer)
     Dim nFilter As Integer
     Screen.MousePointer = 11
     Select Case Index
     Case 1 'Brighter = 14 - (10)          '>0
          nFilter = 14
          EditObjFilter2 nFilter, 10
     Case 2 'Darker = 14 - (-10)           '<0
          nFilter = 14
          EditObjFilter2 nFilter, -10
     Case 3 'Increase Contrast = 17        '>0
          nFilter = 17
          EditObjFilter2 nFilter, 20
     Case 4 'Decrease Contrast = 17        '<0
          nFilter = 17
          EditObjFilter2 nFilter, -20
     Case 5 'Dilate = 25                   '-------
           nFilter = 25
           EditObjFilter2 nFilter, 0
     Case 6 'Erode = 26                    '-------
          nFilter = 26
          EditObjFilter2 nFilter, 0
     Case 7 'Contrast Stretch = 28         '------
          nFilter = 28
          EditObjFilter2 nFilter, 0
     Case 8 'Increase Saturation = 30      '>0
          nFilter = 30
           EditObjFilter2 nFilter, 50
     Case 9 'Decrease Saturation = 30      '<0
          nFilter = 30
           EditObjFilter2 nFilter, -50
     Case Else
          Exit Sub
     End Select
     DrawControl1.Redraw
     Screen.MousePointer = 0
End Sub

Private Sub mnu_Other_Click(Index As Integer)
     Dim nFilter As Integer
     Screen.MousePointer = 11
      
      Select Case Index
       Case 1
            EditObjFilter2 iGrid3d, 20
       Case 2
            EditObjFilter2 iMirrorRL, 0
       Case 3
            EditObjFilter2 iMirrorLR, 0
       Case 4
            EditObjFilter2 iMirrorDT, 0
       Case 5
            EditObjFilter2 iMirrorTD, 0
       Case Else
           Exit Sub
       End Select
     DrawControl1.Redraw
     Screen.MousePointer = 0
End Sub

Private Sub mnuabout_Click()
     frmAbout.ShowForm True
     Unload frmAbout
End Sub

' Move this object to the front of the scene's object list.
Private Sub mnuArrangeSendToBack_Click()
   Screen.MousePointer = 0
   DrawControl1.SetObjectOrder SendToBack
   Screen.MousePointer = 0
End Sub

' Move this object to the Backward of the scene's object list.
Private Sub mnuArrangeSendToBackward_Click()
    Screen.MousePointer = 11
    DrawControl1.SetObjectOrder SendBackward
    Screen.MousePointer = 0
End Sub
' Move this object Send To Forward of the scene's object list.
Private Sub mnuArrangeSendToForward_Click()
   Screen.MousePointer = 11
   DrawControl1.SetObjectOrder BringFoward
   Screen.MousePointer = 0
End Sub

' Move this object Bring To Front of the scene's object list.
Private Sub mnuArrangeSendToFront_Click()
    Screen.MousePointer = 11
    DrawControl1.SetObjectOrder BringToFront
    Screen.MousePointer = 0
End Sub

Private Sub mnuclear_Click()
    Screen.MousePointer = 11
    DrawControl1.ClearObject
    Screen.MousePointer = 0
End Sub

Private Sub MnuCopy_Click()
     Screen.MousePointer = 11
     DrawControl1.CopyObject
     Screen.MousePointer = 0
End Sub

Private Sub mnuCrossHairs_Click()
    DrawControl1.CrossMouse = Not DrawControl1.CrossMouse
    mnuCrossHairs.Checked = DrawControl1.CrossMouse
End Sub

Private Sub mnuCut_Click()
    Screen.MousePointer = 11
    DrawControl1.CutObject
    Screen.MousePointer = 0
End Sub

Private Sub mnuDelete_Click()
   Screen.MousePointer = 11
    DrawControl1.DelObject
    Screen.MousePointer = 0
End Sub

Private Sub mnuEdit_Click()
    Dim mnuenabled1 As Boolean
    Dim mnuenabled2 As Boolean
    
    mnuenabled1 = DrawControl1.IsSelectObject
    mnuenabled2 = FindObject(Clipboard.GetText)

     mnuCut.Enabled = mnuenabled1
     MnuCopy.Enabled = mnuenabled1
     mnuDelete.Enabled = mnuenabled1
     mnupaste.Enabled = mnuenabled2 'True
     mnuclear.Enabled = mnuenabled2 'True
      
    drawToolbar1.EnableButton 6, mnuenabled1 ' mnuCut.Enabled
    drawToolbar1.EnableButton 7, mnuenabled1 'MnuCopy.Enabled
    drawToolbar1.EnableButton 8, mnuenabled2 'mnupaste.Enabled
    drawToolbar1.EnableButton 11, mnuenabled1 ' mnuDelete.Enabled
End Sub

Private Sub mnuEditRedo_Click()
    Redo
End Sub

Private Sub mnuedittext_Click()
      DrawControl1.EditText
End Sub

Private Sub mnuEditUndo_Click()
    Undo
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

' Load the selected file.
Private Sub mnuFileMRU_Click(Index As Integer)
Dim pos As Integer
Dim file_title As String

    If Not DataSafe() Then Exit Sub
    Me.MousePointer = 11
    pos = InStrRev(m_MruList(Index), "\")
    file_title = mid$(m_MruList(Index), pos + 1)
    
    OpenDraw m_MruList(Index), file_title
    
    DrawControl1.Redraw
    Me.MousePointer = 0
End Sub

'Open file draw
Sub OpenDraw(Filename As String, FileTitle As String)
    Screen.MousePointer = 11
    If DrawControl1.OpenDraw(Filename, FileTitle) Then
      ' Update the caption.
       SetFileName Filename, FileTitle
      
      ' DrawControl1.Redraw
       
    End If
    Screen.MousePointer = 0
End Sub

' Start a new picture.
Private Sub mnuFileNew_Click()
    If Not DataSafe() Then Exit Sub

    'New draw
    DrawControl1.NewDraw True
     
    'Blank the file name.
    SetFileName "", ""

    'The data has not been modified.
    m_DataModified = False
    DrawControl1.SetScaleFull
    ' Prepare to edit.
    DrawControl1.PrepareToEdit
End Sub

' Load a file.
Private Sub mnuFileOpen_Click()
Dim File_name As String
Dim sOpen As SelectedFile

   FileDialog.sFilter = "ArtDraw Files (*.adrw)" & Chr$(0) & "*.adrw" '& Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
   FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
   FileDialog.sInitDir = App.Path & "\Samples"
   sOpen = ShowOpen(Me.hWnd)
   If Err.Number <> 32755 And sOpen.bCanceled = False Then
      File_name = sOpen.sFile '(1)
      Screen.MousePointer = 11
      OpenDraw File_name, File_name
      
      Screen.MousePointer = 0
    End If
   
End Sub

' Save the data using the current file name.
Private Sub mnuFileSave_Click()
    If Len(DrawControl1.Filename) = 0 Then
        ' There is no file name. Use Save As.
        mnuFileSaveAs_Click
    Else
        ' Save the data.
        If DrawControl1.SaveDraw(DrawControl1.Filename, DrawControl1.FileTitle) Then
           Screen.MousePointer = 11
           ' Update the caption.
           SetFileName DrawControl1.Filename, DrawControl1.FileTitle
           Screen.MousePointer = 0
        End If
    End If
End Sub

' Save the picture with a new file name.
Private Sub mnuFileSaveAs_Click()
Dim File_name As String
Dim sSave As SelectedFile

    FileDialog.sFilter = "ArtDraw Files (*.adrw)" & Chr$(0) & "*.adrw" '& Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_HIDEREADONLY
    FileDialog.sInitDir = App.Path & "\Samples"
    FileDialog.sDefFileExt = "*.adrw"
    sSave = ShowSave(Me.hWnd)
    If Err.Number <> 32755 And sSave.bCanceled = False Then
        Screen.MousePointer = 11
        File_name = sSave.sFile
        If DrawControl1.SaveDraw(File_name, File_name) Then
          ' Update the caption.
          SetFileName DrawControl1.Filename, DrawControl1.FileTitle
       End If
       Screen.MousePointer = 0
    End If
    
'    dlgFile.flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNOverwritePrompt
'
'    If PathExists(App.Path + "\Samples") = False Then MkDir App.Path + "\Samples"
'    dlgFile.InitDir = App.Path + "\Samples"
'    dlgFile.Filename = DrawControl1.Filename
'    dlgFile.Filter = "ArtDraw Files (*.adrw)|*.adrw|" & "All Files (*.*)|*.*"
'    On Error Resume Next
'    dlgFile.ShowSave
'    If Err.Number = cdlCancel Then
'        Exit Sub
'    ElseIf Err.Number <> 0 Then
'        MsgBox "Error " & Format$(Err.Number) & " selecting file." & vbCrLf & Err.Description
'        Exit Sub
'    End If
'
'    File_name = dlgFile.Filename
'    dlgFile.InitDir = left$(File_name, Len(File_name) - Len(dlgFile.FileTitle) - 1)
'   Screen.MousePointer = 11
'   If DrawControl1.SaveDraw(File_name, dlgFile.FileTitle) Then
'     ' Update the caption.
'      SetFileName DrawControl1.Filename, DrawControl1.FileTitle
'   End If
'   Screen.MousePointer = 0
End Sub

' Save a bitmap image.
Private Sub mnuFileSaveBitmap_Click()
Dim old_file_name As String
Dim pos As Integer
Dim File_name As String
Dim sSave As SelectedFile
Dim fDrive As String, fPath As String, fFileName As String, fFile As String, fExtension As String
                    
    FileDialog.sFile = ""
    FileDialog.sFilter = "Bitmap Files (*.bmp)" + Chr(0) + "*.bmp" + Chr(0) + "" + _
                         "Graphics Interchange Format(*.gif)" + Chr(0) + "*.gif" + Chr(0) + "" + _
                         "Tagged Image Format(*.tif)" + Chr(0) + "*.tif" + Chr(0) + "" + _
                         "Portable Network Graphics(*.png)" + Chr(0) + "*.png" + Chr(0) + "" + _
                         "Joint Photographic Experts Group(*.jpg)" + Chr(0) + "*.jpg" + Chr(0) + _
                         "Metafiles (*.wmf)" + Chr(0) + "*.wmf" + Chr(0) + _
                         "Ascii file CNC Pagra(*.txt)" + Chr(0) + "*.txt"
                           
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_HIDEREADONLY
    FileDialog.sInitDir = App.Path & "\Export"
    FileDialog.sDefFileExt = "*.bmp"
    sSave = ShowSave(Me.hWnd)
    If Err.Number <> 32755 And sSave.bCanceled = False Then
        Screen.MousePointer = 11
        File_name = sSave.sFile
        
        'dlgFile.InitDir = left$(File_name, Len(File_name) - Len(dlgFile.FileTitle) - 1)
        SplitPath File_name, fDrive, fPath, fFileName, fFile, fExtension
        Screen.MousePointer = 11
        Select Case FileDialog.nFilterIndex
        Case 1 'bmp
             fExtension = ".bmp"
        Case 2 'gif
             fExtension = ".gif"
        Case 3 'tif
             fExtension = ".tif"
        Case 4 'png
             fExtension = ".png"
        Case 5 'jpg
             fExtension = ".jpg"
        Case 6 'wmf
             fExtension = ".wmf"
        Case 7 '
             fExtension = ".txt"
             ExportOn True
        Case Else
             MsgBox "Error Extension " & File_name & " Not Saved!"
            Exit Sub
        End Select
        
        If Left(fExtension, 1) <> "." Then fExtension = "." + fExtension
        File_name = fPath + "\" + fFile + fExtension
    
        If FileDialog.nFilterIndex = 6 Then
            DrawControl1.FileExport File_name
        ElseIf FileDialog.nFilterIndex = 7 Then
            DrawControl1.Redraw
            
            SaveExportTxt Me, File_name
             ExportOn False
             If FileExists(File_name) Then
                MsgBox File_name & " Saved OK!"
             Else
                MsgBox "ERROR! " & File_name & " Not Saved!"
             End If
        Else
            DrawControl1.FileExportBitmap File_name
        End If
        
        
        'dlgFile.Filename = old_file_name
        Screen.MousePointer = 0
        Screen.MousePointer = 0
    End If
                     
'    old_file_name = dlgFile.Filename
'    pos = InStrRev(old_file_name, ".")
'    If pos > 0 Then dlgFile.Filename = left$(old_file_name, pos) & "bmp"
'
'    dlgFile.flags = cdlOFNExplorer Or _
'        cdlOFNHideReadOnly Or _
'        cdlOFNLongNames Or _
'        cdlOFNOverwritePrompt
'    dlgFile.Filter = "Bitmap Files (*.bmp)|*.bmp|" + _
'                     "Graphics Interchange Format(*.gif)|*.gif|" + _
'                     "Tagged Image Format(*.tif)|*.tif|" + _
'                     "Portable Network Graphics(*.png)|*.png|" + _
'                     "Joint Photographic Experts Group(*.jpg)|*.jpg|Metafiles (*.wmf)|*.wmf|" & _
'                     "All Files (*.*)|*.*"
'
'    If PathExists(App.Path + "\Export") = False Then MkDir App.Path + "\Export"
'    dlgFile.InitDir = App.Path + "\Export"
'    On Error Resume Next
'    dlgFile.ShowSave
'    If Err.Number = cdlCancel Then
'        Exit Sub
'    ElseIf Err.Number <> 0 Then
'        MsgBox "Error " & Format$(Err.Number) & _
'            " selecting file." & vbCrLf & _
'            Err.Description
'        Exit Sub
'    End If

'    File_name = dlgFile.Filename
'    dlgFile.InitDir = left$(File_name, Len(File_name) - Len(dlgFile.FileTitle) - 1)
'    SplitPath File_name, fDrive, fPath, fFileName, fFile, fExtension
'    Screen.MousePointer = 11
'    Select Case dlgFile.FilterIndex
'    Case 1 'bmp
'         fExtension = ".bmp"
'    Case 2 'gif
'         fExtension = ".gif"
'    Case 3 'tif
'         fExtension = ".tif"
'    Case 4 'png
'         fExtension = ".png"
'    Case 5 'jpg
'         fExtension = ".jpg"
'    Case 6 'wmf
'         fExtension = ".wmf"
'    Case Else
'         MsgBox "Error Extension " & File_name & " Not Saved!"
'         Exit Sub
'    End Select
'    If left(fExtension, 1) <> "." Then fExtension = "." + fExtension
'    File_name = fPath + "\" + fFile + fExtension
'
'    If dlgFile.FilterIndex = 6 Then
'       DrawControl1.FileExport File_name
'    Else
'       DrawControl1.FileExportBitmap File_name
'    End If
'
'    dlgFile.Filename = old_file_name
'    Screen.MousePointer = 0
End Sub

' Save the objects in a metafile.
Private Sub mnuFileSaveMetafile_Click()
Dim old_file_name As String
Dim pos As Integer
Dim File_name As String

'   ' old_file_name = dlgFile.Filename
'    pos = InStrRev(old_file_name, ".")
'    If pos > 0 Then dlgFile.Filename = Left$(old_file_name, pos) & "wmf"
'
'    dlgFile.flags = cdlOFNExplorer Or _
'        cdlOFNHideReadOnly Or _
'        cdlOFNLongNames Or _
'        cdlOFNOverwritePrompt
'    dlgFile.Filter = "Metafiles (*.wmf)|*.wmf|" & _
'        "All Files (*.*)|*.*"
'    On Error Resume Next
'    dlgFile.ShowSave
'    If Err.Number = cdlCancel Then
'        Exit Sub
'    ElseIf Err.Number <> 0 Then
'        MsgBox "Error " & Format$(Err.Number) & _
'            " selecting file." & vbCrLf & _
'            Err.Description
'        Exit Sub
'    End If
'
'    File_name = dlgFile.Filename
'    dlgFile.InitDir = Left$(File_name, Len(File_name) _
'        - Len(dlgFile.FileTitle) - 1)

'    dlgFile.Filename = old_file_name
End Sub


Private Sub mnufillform_Click()
       Open_Form 13 '"Fill"
End Sub

Private Sub mnuFilter_Click()
     EditObjFilter
End Sub

Private Sub mnuImport_Click()
    Dim old_file_name As String
    Dim pos As Integer
    Dim File_name As String
    Dim sSave As SelectedFile
    Dim fDrive As String, fPath As String, fFileName As String, fFile As String, fExtension As String
                     
    FileDialog.sFilter = "Bitmap Files (*.bmp,*.gif,*.tif,*.png,*.jpg,*.wmf)" + Chr(0) + "*.bmp;*.gif;*.tif;*.png;*.jpg;*.wmf;*.emf"
                           
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_HIDEREADONLY
    FileDialog.sInitDir = App.Path & "\Object"
    FileDialog.sDefFileExt = "*.bmp"
    sSave = ShowOpen(Me.hWnd)
    If Err.Number <> 32755 And sSave.bCanceled = False Then
       File_name = sSave.sLastDirectory + sSave.sFile
       'dlgFile.InitDir = Left$(File_name, Len(File_name) _
        - Len(dlgFile.FileTitle) - 1)

       'LoadPicBox File_name,
        DrawControl1.Pattern = File_name
        DrawControl1.SelectTool 15
    End If
'    dlgFile.flags = cdlOFNExplorer Or _
'                    cdlOFNHideReadOnly Or _
'                    cdlOFNLongNames Or _
'                    cdlOFNOverwritePrompt
'    dlgFile.Filter = "Bitmap Files (bmp,gif,tif,png,jpg,wmf)|*.bmp;*.gif;*.tif;*.png;*.jpg;*.wmf"
'
'    If PathExists(App.Path + "\Object") = False Then MkDir App.Path + "\Object"
'    dlgFile.InitDir = App.Path + "\Object"
'    On Error Resume Next
'    dlgFile.ShowOpen
'    If Err.Number = cdlCancel Then
'        Exit Sub
'    ElseIf Err.Number <> 0 Then
'        MsgBox "Error " & Format$(Err.Number) & _
'            " selecting file." & vbCrLf & _
'            Err.Description
'        Exit Sub
'    End If

'    File_name = dlgFile.Filename
'    dlgFile.InitDir = Left$(File_name, Len(File_name) _
'        - Len(dlgFile.FileTitle) - 1)
'
'    'LoadPicBox File_name,
'     DrawControl1.Pattern = File_name
'     DrawControl1.SelectTool 15
End Sub

Private Sub mnuMove_Click()
      DrawControl1.ViewTransform 0
End Sub

Private Sub Mnunormal_Click()
     m_ViewSimple = False
     Mnunormal.Checked = True
     mnuSimpleWireframe.Checked = False
     DrawControl1.Redraw
End Sub

Private Sub mnuObjectPoint_Click()
     DrawControl1.ShowObjectPoint = True
End Sub

Private Sub mnuPaste_Click()
     Screen.MousePointer = 11
     DrawControl1.PasteObject
     Screen.MousePointer = 0
End Sub

Private Sub mnupenform_Click()
     Open_Form 12 '"Pen"
End Sub

Private Sub mnuPrint_Click()
     DrawControl1.PrintDraw
End Sub

Private Sub mnuprintersetup_Click()
    Dim PSetup As PrinterSetup
   
    If Printers.Count > 0 Then
     If gPrintetOrientation > 0 Then
        Printer.Orientation = gPrintetOrientation
     End If
     PSetup = ShowPageSetupDlg(Me.hWnd)
     If PSetup.Printer.dmOrientation > 0 Then
        gPrintetOrientation = PSetup.Printer.dmOrientation
        Printer.Orientation = gPrintetOrientation
     End If
    End If
  
End Sub

Private Sub mnuReflect_Click()
     DrawControl1.ViewTransform 4
End Sub

Private Sub mnuRuler_Click()
      DrawControl1.DrawRuler = Not DrawControl1.DrawRuler
      mnuRuler.Checked = DrawControl1.DrawRuler
End Sub

Private Sub mnuSimpleWireframe_Click()
     m_ViewSimple = True
     mnuSimpleWireframe.Checked = True
     Mnunormal.Checked = False
     DrawControl1.Redraw
End Sub

Private Sub mnuskew_Click()
     DrawControl1.ViewTransform 3
End Sub

Private Sub mnuSymbol_Click()
     m_FormSymbolView = True
     FrmSymbols.Show
End Sub

Private Sub mnuDownLoad1_Click(Index As Integer)
     Select Case Index
     Case 1: Execute "http://freetextures.org/"
     Case 2: Execute "http://www.free-pictures-photos.com/index.htm"
     Case 3: Execute "http://www.squidfingers.com/patterns/1/"
     Case 4: Execute "http://www.pickafont.com/"
     Case 5: Execute "http://www.webpagepublicity.com/free-fonts-c.html"
     Case 6: Execute "http://www.1001freefonts.com/"
     End Select
End Sub

' Clear the selected objects' transformations.
Private Sub mnuTransformClear_Click()
    DrawControl1.ClearTransform
End Sub

Private Sub mnutransformform_Click()
      DrawControl1.ViewTransform 0
End Sub

' Rotate the selected objects.
Private Sub mnuTransformRotate_Click()
    DrawControl1.ViewTransform 1
   
End Sub

' Let the user scale the selected objects.
Private Sub mnuTransformScale_Click()
 
     DrawControl1.ViewTransform 2
     
End Sub


Private Sub InitScreen()
Dim i As Integer
Dim FileExt As New clsExtReg
Dim ret As Boolean
    
    'Scalemode page
    gScaleMode = GetSetting(App.ProductName, "Main", "ScaleMode", vbPixels)
    
    ComboZoom.AddItem "10 %"
    ComboZoom.AddItem "25 %"
    ComboZoom.AddItem "50 %"
    ComboZoom.AddItem "100 %"
    ComboZoom.AddItem "150 %"
    ComboZoom.AddItem "200 %"
    ComboZoom.AddItem "400 %"

    ColorPalette1_ColorSelected 1, 16777215
    ColorPalette1_ColorSelected 2, 0
    
    drawToolbar.BarOrientation = tbVertical
    drawToolbar.BuildToolbar PicTools.Picture, vbButtonFace, 16, "OOOOOOOOOOOOO"
    drawToolbar.SetTooltips "Arrow|Point|Polyline|FreePolygon|Free Line|Calligraphy|Curve|RectAngle|Polygon|Ellipse|Text Art|Pen|Fill"
    drawToolbar.CheckButton 1, True
    
    drawToolbar2.BarOrientation = tbHorizontal
    drawToolbar2.BuildToolbar PicTollBar2.Picture, vbButtonFace, 16, "NN|NNNN|C|NN"
    drawToolbar2.SetTooltips "Select all|UnselectAll|Bring to front|Send to back|Bring Forward|Send Backward|Lock|Group|Ungroup"
    For i = 3 To 8
       If i <> 7 Then
       drawToolbar2.EnableButton i, False
       End If
    Next
    
    drawToolbar1.BarOrientation = tbHorizontal
    drawToolbar1.BuildToolbar PicTollBar1.Picture, vbButtonFace, 16, "NNNNN|NNN|NN|N|NNNN"
    drawToolbar1.SetTooltips "New draw|Open draw|Save draw|Export draw|Print|Cut|Copy|Paste|Undo|Redo|Delete|Symbol|Pen|Fill|Transforming"
    For i = 6 To 11 '2
       drawToolbar1.EnableButton i, False
    Next
    
    drawToolbar3.BarOrientation = tbHorizontal
    drawToolbar3.BuildToolbar PicTollBar3.Picture, vbButtonFace, 16, "NNNNNNNNN"
    drawToolbar3.SetTooltips "Zoom Windows|Zoom All|Zoom (-)|Zoom (+)|Zoom object (Ctl+F1)||||Pan"
    
'    drawToolbar.BarEdge = True
'    drawToolbar1.BarEdge = True
'    drawToolbar2.BarEdge = True
'    drawToolbar3.BarEdge = True
'
    drawToolbar2.EnableButton 8, False
    drawToolbar2.EnableButton 9, False
    ' CoolBar2.Width = 500
    'register .adrw
    If FileExt.GetBinaryValue("HKEY_CLASSES_ROOT\.adrw", "") <> App.EXEName + ".exe" Then
       ret = FileExt.Register(".adrw", "Art Draw", App.Path & "\" + App.EXEName + ".exe", True, App.Path & "\" + App.EXEName + ".exe")
    End If
    
    mnuText.Enabled = False
    mnuBitMap.Enabled = False
    
    ' Load the MRU list.
    MruLoad
        
    gPrintetOrientation = Printer.Orientation
        
    ' Start a new picture.
    Me.Width = Screen.Width - 1000
    Me.Height = Screen.Height - 1000
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    DrawControl1.CrossMouse = GetSetting(App.ProductName, "MOUSE", "CROSS", Trim(Str(False)))
    mnuCrossHairs.Checked = DrawControl1.CrossMouse
    
    DrawControl1.DrawRuler = GetSetting(App.ProductName, "DRAWCONTROL", "DrawRuler", Trim(Str(False)))
    mnuRuler.Checked = DrawControl1.DrawRuler
      
End Sub
   
' Select the arrow tool.
Public Sub SelectArrowTool()

    ' Make sure the arrow button is pressed.
    drawToolbar.CheckButton 1, True
    ' Prepare to deal with this tool.
    DrawControl1.SelectTool 1 '"Arrow"
End Sub

Sub Open_Form(Index As Integer)
    DrawControl1.SelectTool Index
    
'    If Index = 12 Or Index = 13 Then
'       SelectArrowTool
'    End If
End Sub

Private Sub EditObjFilter()
    Dim fPic As StdPicture
     Set fPic = DrawControl1.ObjPicture
     If FrmFilter.ShowForm(fPic) = False Then
           Set DrawControl1.ObjPicture = fPic
           DrawControl1.Redraw
     End If
End Sub


Private Sub EditObjFilter2(Optional nFilter As Integer = 0, Optional Factor As Long = 0)
     Dim fPic As StdPicture, pProgress As Long
   
     Set fPic = DrawControl1.ObjPicture
     FilterG nFilter, fPic, Factor, pProgress
     Set DrawControl1.ObjPicture = fPic
   
End Sub

Public Function ShowColor(mColor As Long) As Boolean
     Dim C As SelectedColor
    
    ColorDialog.rgbResult = mColor
    C = CommonDialog.ShowColor(Me.hWnd, False)
    ShowColor = Not C.bCanceled
    If C.bCanceled = False Then
      mColor = C.oSelectedColor
    End If

End Function

Public Function ScalePage(gSM As Integer) As String
      Select Case gSM
      Case vbPixels
           ScalePage = "(Pix)"
      Case vbInches
           ScalePage = "(In)"
      Case vbMillimeters
           ScalePage = "(mm)"
      Case Else
         Exit Function
      End Select
End Function

Public Sub WheelMoved(ByVal delta As Long, X As Long, Y As Long)
        'Debug.Print delta, X, Y
End Sub
