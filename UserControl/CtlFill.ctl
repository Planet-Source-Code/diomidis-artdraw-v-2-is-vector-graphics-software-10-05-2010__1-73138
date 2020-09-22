VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl CtlFill 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9240
   KeyPreview      =   -1  'True
   ScaleHeight     =   2835
   ScaleWidth      =   9240
   ToolboxBitmap   =   "CtlFill.ctx":0000
   Begin VB.OptionButton Option4 
      Height          =   345
      Left            =   1230
      Picture         =   "CtlFill.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Fill image"
      Top             =   75
      Width           =   375
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1900
      Left            =   6705
      ScaleHeight     =   1905
      ScaleWidth      =   2175
      TabIndex        =   30
      Top             =   525
      Visible         =   0   'False
      Width           =   2175
      Begin VB.CommandButton Command3 
         Caption         =   "Apply"
         Height          =   315
         Left            =   570
         TabIndex        =   32
         Top             =   1560
         Width           =   1485
      End
      Begin VB.PictureBox mPicture 
         Height          =   480
         Left            =   135
         ScaleHeight     =   420
         ScaleWidth      =   450
         TabIndex        =   35
         Top             =   555
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.TextBox TextFile 
         Height          =   360
         Left            =   45
         TabIndex        =   34
         Top             =   45
         Width           =   1650
      End
      Begin VB.CommandButton CommandPicture 
         Caption         =   "..."
         Height          =   345
         Left            =   1710
         TabIndex        =   33
         Top             =   60
         Width           =   360
      End
      Begin VB.PictureBox Picture8 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   120
         Left            =   210
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   31
         Top             =   1590
         Width           =   120
      End
      Begin VB.Image ImagePic 
         Height          =   975
         Left            =   105
         Stretch         =   -1  'True
         Top             =   525
         Width           =   1965
      End
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      LargeChange     =   10
      Left            =   45
      Max             =   254
      TabIndex        =   28
      Top             =   2625
      Width           =   1900
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   255
      Pattern         =   "*.bmp"
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.OptionButton Option3 
      Height          =   345
      Left            =   870
      Picture         =   "CtlFill.ctx":0654
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Fill Gradient"
      Top             =   75
      Width           =   375
   End
   Begin VB.OptionButton Option2 
      Height          =   345
      Left            =   540
      Picture         =   "CtlFill.ctx":0A0A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Fill pattern"
      Top             =   75
      Width           =   330
   End
   Begin VB.OptionButton Option1 
      Height          =   345
      Left            =   150
      Picture         =   "CtlFill.ctx":0DC0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Fill color"
      Top             =   75
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1900
      Left            =   4500
      ScaleHeight     =   1905
      ScaleWidth      =   2100
      TabIndex        =   8
      Top             =   510
      Visible         =   0   'False
      Width           =   2100
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   45
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   45
         Width           =   1950
      End
      Begin VB.PictureBox PictureGrd2 
         BackColor       =   &H00000000&
         Height          =   450
         Left            =   1200
         ScaleHeight     =   390
         ScaleWidth      =   405
         TabIndex        =   23
         Top             =   900
         Width           =   465
      End
      Begin VB.PictureBox PictureGrd1 
         BackColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   1200
         ScaleHeight     =   330
         ScaleWidth      =   390
         TabIndex        =   22
         Top             =   510
         Width           =   450
      End
      Begin VB.PictureBox PictureGradient 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1100
         Left            =   45
         ScaleHeight     =   69
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   69
         TabIndex        =   21
         Top             =   405
         Width           =   1100
      End
      Begin VB.CommandButton CommandGrd 
         Caption         =   "Apply"
         Height          =   315
         Left            =   570
         TabIndex        =   20
         Top             =   1560
         Width           =   1485
      End
      Begin VB.CommandButton CommandGrd1 
         Height          =   420
         Left            =   1605
         Picture         =   "CtlFill.ctx":1176
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   495
         Width           =   435
      End
      Begin VB.CommandButton CommandGrd2 
         Height          =   420
         Left            =   1620
         Picture         =   "CtlFill.ctx":18E0
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   915
         Width           =   435
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1900
      Left            =   2355
      ScaleHeight     =   1905
      ScaleWidth      =   2175
      TabIndex        =   7
      Top             =   495
      Visible         =   0   'False
      Width           =   2175
      Begin VB.PictureBox PicturePattern 
         BackColor       =   &H00FFFFFF&
         Height          =   1100
         Left            =   60
         ScaleHeight     =   1035
         ScaleWidth      =   1905
         TabIndex        =   13
         Top             =   405
         Width           =   1965
      End
      Begin VB.PictureBox PicturePat1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1275
         ScaleHeight     =   300
         ScaleWidth      =   315
         TabIndex        =   26
         Top             =   510
         Width           =   375
      End
      Begin VB.PictureBox PicturePat2 
         BackColor       =   &H00000000&
         Height          =   360
         Left            =   1275
         ScaleHeight     =   300
         ScaleWidth      =   330
         TabIndex        =   25
         Top             =   900
         Width           =   390
      End
      Begin VB.PictureBox Image1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   120
         Left            =   210
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   17
         Top             =   1590
         Width           =   120
      End
      Begin VB.CommandButton CommandColor2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1665
         TabIndex        =   16
         Top             =   900
         Width           =   360
      End
      Begin VB.CommandButton CommandColor1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1665
         TabIndex        =   15
         Top             =   525
         Width           =   360
      End
      Begin VB.CommandButton CommandPattern 
         Caption         =   "Apply"
         Height          =   315
         Left            =   570
         TabIndex        =   14
         Top             =   1560
         Width           =   1485
      End
      Begin VB.ComboBox List1 
         Height          =   315
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   45
         Width           =   1965
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1905
      Left            =   45
      ScaleHeight     =   1905
      ScaleWidth      =   2100
      TabIndex        =   0
      Top             =   465
      Width           =   2100
      Begin VB.CommandButton CommandFill 
         Caption         =   "Apply"
         Height          =   315
         Left            =   570
         TabIndex        =   4
         Top             =   1560
         Width           =   1485
      End
      Begin VB.PictureBox PicFillColor 
         BackColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   105
         ScaleHeight     =   405
         ScaleWidth      =   1320
         TabIndex        =   3
         Top             =   285
         Width           =   1380
      End
      Begin VB.CommandButton cmdSysColorsFill 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Left            =   1485
         Picture         =   "CtlFill.ctx":204A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   " System colors "
         Top             =   240
         Width           =   450
      End
      Begin MSComctlLib.ImageList imlFillStyles 
         Left            =   1425
         Top             =   780
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CtlFill.ctx":27B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CtlFill.ctx":29C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CtlFill.ctx":2BD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CtlFill.ctx":2DEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CtlFill.ctx":2FFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CtlFill.ctx":320E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CtlFill.ctx":3420
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CtlFill.ctx":3632
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageCombo icbFillStyle 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "imlFillStyles"
      End
      Begin VB.Label LblFillStyle 
         BackStyle       =   0  'Transparent
         Caption         =   "Fill Style"
         Height          =   240
         Left            =   150
         TabIndex        =   6
         Top             =   840
         Width           =   1785
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fill Color"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   60
         Width           =   1800
      End
   End
   Begin VB.Label Labelblend 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blend: 0%"
      Height          =   195
      Left            =   45
      TabIndex        =   29
      Top             =   2370
      Width           =   705
   End
End
Attribute VB_Name = "CtlFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(c) 2007 - Diomidisk

'Default Property Values:
Const m_def_Blend = 0
Const m_def_TypeGradient = 0
Const m_def_TypeFill = 1
Const m_def_Color1 = 0
Const m_def_Color2 = 0
Const m_def_FillStyle = 0
Const m_def_NamePattern = ""
'Property Variables:
Dim m_Blend As Integer
Dim m_TypeGradient As Integer
Dim m_TypeFill As Integer
Dim m_Color1 As OLE_COLOR
Dim m_Color2 As OLE_COLOR
Dim m_FillStyle As Integer
Dim m_NamePattern As String

Dim PicPatten() As Byte
Dim ImageData(7, 7) As Long
Dim nColor() As Long

Event Apply(nTypeFill As Integer, nFillStyle As Integer, nColor1 As Long, nColor2 As Long, nPattern As String, nTypeGradient As Integer, mBlend As Integer)
Event ApplyImage(nTypeFill As Integer, nFillStyle As Integer, nPattern As String, nPicture As StdPicture, mBlend As Integer)

Private Sub cmdSysColorsFill_Click()
     OpenColorDialog PicFillColor
End Sub

Private Sub CommandPicture_Click()
Dim File_name As String
Dim sOpen As SelectedFile
                     
    FileDialog.sFilter = "Image Files" + Chr(0) + "*.bmp;*.gif;*.tif;*.png;*.jpg;*.wmf;*.emf"
                           
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_HIDEREADONLY
    FileDialog.sInitDir = App.Path & "\Object"
    'FileDialog.sDefFileExt = "*.jpg"
    sOpen = ShowOpen(UserControl.hWnd, False)
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
       File_name = sOpen.sFile
       ReadPicturePattern File_name
    End If
'    dlgFile.flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNFileMustExist
'    dlgFile.Filter = "Bitmap Files |*.bmp;*.jpg;*.gif"
'    If PathExists(App.Path + "\Object") = False Then MkDir App.Path + "\Object"
'    dlgFile.InitDir = App.Path + "\Object"
'    On Error Resume Next
'    dlgFile.ShowOpen
'    If Err.Number = cdlCancel Then
'        Exit Sub
'    ElseIf Err.Number <> 0 Then
'        MsgBox "Error " & Format$(Err.Number) & " selecting file." & vbCrLf & Err.Description
'        Exit Sub
'    End If

   ' File_name = dlgFile.Filename
   ' ReadPicturePattern File_name
    
    
End Sub

Private Sub ReadPicturePattern(File_name As String)
      Dim Bitmap As Long
      TextFile = File_name
      
      'LoadPicBox File_name, mPicture
      
      mPicture = LoadPictureGDIPlus(File_name) 'LoadPicture(File_name)
      
      mPicture.AutoSize = True
      'LoadPicBox File_name, mPicture.hDC, , Bitmap
      ImagePic.Picture = mPicture.Image
      'ImagePic.Picture = Bitmap
End Sub
Private Sub Combo1_Click()
    If Combo1.ListIndex = -1 Then Exit Sub
    TypeGradient = Combo1.ListIndex
    'DrawGradient
   ' PictureGradient.Refresh
End Sub

Private Sub Command3_Click()
      If FileExists(TextFile.Text) Then
         TypeFill = 4
         FillStyle = 10
         NamePattern = TextFile
         RaiseEvent ApplyImage(TypeFill, FillStyle, NamePattern, mPicture, Blend)
      End If
End Sub

Private Sub CommandColor1_Click()
    Dim Oldcolor As Long
    Oldcolor = PicturePat1.BackColor
    OpenColorDialog PicturePat1
    If Oldcolor <> PicturePat1.BackColor Then
       ReplaceColor nColor(1), PicturePat1.BackColor
       Color1 = PicturePat1.BackColor
       TileIt PicturePattern, Image1
       nColor(1) = Color1
    End If
End Sub

Private Sub CommandColor2_Click()
    Dim Oldcolor As Long
    
    Oldcolor = PicturePat2.BackColor
    OpenColorDialog PicturePat2
    If Oldcolor <> PicturePat2.BackColor Then
       ReplaceColor nColor(2), PicturePat2.BackColor
       Color2 = PicturePat2.BackColor
       TileIt PicturePattern, Image1
       nColor(2) = Color2
    End If
End Sub

Private Sub CommandFill_Click()
      TypeFill = 1
      FillStyle = icbFillStyle.SelectedItem.Index - 1
      Color1 = PicFillColor.BackColor
      Color2 = 0
      NamePattern = ""
      mRaiseEvent
      
End Sub

Private Sub CommandGrd_Click()
      TypeGradient = Combo1.ListIndex
      TypeFill = 3
      FillStyle = 9
      mRaiseEvent
End Sub

Private Sub CommandGrd1_Click()
    OpenColorDialog PictureGrd1
    If Color1 <> PictureGrd1.BackColor Then
       Color1 = PictureGrd1.BackColor
       DrawGradient
    End If
End Sub

Private Sub CommandGrd2_Click()
    OpenColorDialog PictureGrd2
    If Color2 <> PictureGrd2.BackColor Then
       Color2 = PictureGrd2.BackColor
       DrawGradient
    End If
End Sub

Private Sub CommandPattern_Click()
      TypeFill = 2
      FillStyle = 8
      NamePattern = File1.Filename
      mRaiseEvent
End Sub

Private Sub HScroll2_Change()
       m_Blend = HScroll2.Value
       Labelblend.Caption = "Blend:" + Format(100 * (m_Blend / 254), "0.0") + "%"
End Sub

Private Sub icbFillStyle_Click()
     icbFillStyle.ToolTipText = icbFillStyle.SelectedItem.Key
End Sub



Private Sub Option1_Click()
       Picture1.Visible = True
       Picture2.Visible = False
       Picture3.Visible = False
       Picture4.Visible = False
End Sub

Private Sub Option2_Click()
      Picture1.Visible = False
      Picture2.Visible = True
      Picture3.Visible = False
      Picture4.Visible = False
End Sub

Private Sub Option3_Click()
     Picture1.Visible = False
     Picture2.Visible = False
     Picture3.Visible = True
     Picture4.Visible = False
End Sub

Private Sub List1_Click()
       File1.ListIndex = List1.ListIndex
       Image1.Picture = LoadPicture(File1.Path + "\" + File1.Filename)
       ReadImagedata
       TileIt PicturePattern, Image1
       If UBound(nColor) > 0 Then
            ReplaceColor nColor(1), PicturePat1.BackColor
            nColor(1) = PicturePat1.BackColor
            ReplaceColor nColor(2), PicturePat2.BackColor
            nColor(2) = PicturePat2.BackColor
       End If
End Sub




Private Sub Option4_Click()
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = True
End Sub

Private Sub PicFillColor_DblClick()
cmdSysColorsFill_Click
End Sub

Private Sub PictureGrd1_DblClick()
CommandGrd1_Click
End Sub



Private Sub PictureGrd2_DblClick()
CommandGrd2_Click
End Sub

Private Sub UserControl_Initialize()
     Picture2.Move Picture1.Left, Picture1.Top, Picture1.Width, Picture1.Height
     Picture3.Move Picture1.Left, Picture1.Top, Picture1.Width, Picture1.Height
     Picture4.Move Picture1.Left, Picture1.Top, Picture1.Width, Picture1.Height
     Labelblend.Move Picture1.Left, Picture1.Height + Picture1.Top
     HScroll2.Move Picture1.Left ', Picture1.Height + Picture1.Top + LabelBlend.Height + LabelBlend.Top
     HScroll2.Width = Picture1.Width
     
     Picture2.Visible = False
     Picture3.Visible = False
     Picture4.Visible = False
     If PathExists(App.Path + "\Pattern") = False Then MkDir App.Path + "\Pattern"
     File1.Path = App.Path + "\Pattern"
     m_Color1 = RGB(255, 255, 255)
     m_Color2 = RGB(0, 0, 0)
     Color1 = RGB(255, 255, 255)
     Color2 = RGB(0, 0, 0)
     FillPatternList
     DrawFill
End Sub

Sub TileIt(Obj As Object, ImageSource As Object)

    Dim i As Integer, j As Integer
    Dim Result1 As Integer, Result2 As Integer
    Obj.AutoRedraw = True ' Set the form autoredraw
    Obj.Refresh ' and refresh
    Result1 = Obj.Height / ImageSource.Height
    Result2 = Obj.Width / ImageSource.Width
    Image1.Picture = Image1.Image
    For i = 0 To Result1
        For j = 0 To Result2
            Obj.PaintPicture ImageSource.Picture, j * _
            ImageSource.Width, i * ImageSource.Height, _
            ImageSource.Width, ImageSource.Height
        Next
    Next

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     
    
    m_Color1 = PropBag.ReadProperty("Color1", m_def_Color1)
    m_Color2 = PropBag.ReadProperty("Color2", m_def_Color2)
    m_FillStyle = PropBag.ReadProperty("FillStyle", m_def_FillStyle)
    m_NamePattern = PropBag.ReadProperty("NamePattern", m_def_NamePattern)
    m_TypeFill = PropBag.ReadProperty("TypeFill", m_def_TypeFill)
    Picture1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Picture2.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Picture3.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Picture4.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_TypeGradient = PropBag.ReadProperty("TypeGradient", m_def_TypeGradient)
    
    Picture2.Move Picture1.Left, Picture1.Top, Picture1.Width, Picture1.Height
    Picture3.Move Picture1.Left, Picture1.Top, Picture1.Width, Picture1.Height
    Picture4.Move Picture1.Left, Picture1.Top, Picture1.Width, Picture1.Height
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False
    CommandFill.Move CommandFill.Left, CommandFill.Top, CommandFill.Width, CommandFill.Height
    CommandGrd.Move CommandFill.Left, CommandFill.Top, CommandFill.Width, CommandFill.Height
    m_Color1 = RGB(255, 255, 255)
    m_Color2 = RGB(0, 0, 0)
     
    If PathExists(App.Path + "\Pattern") = False Then MkDir App.Path + "\Pattern"
    File1.Path = App.Path + "\Pattern"
    FillPatternList
    DrawFill
    
    m_Blend = PropBag.ReadProperty("Blend", m_def_Blend)
End Sub

Private Sub FillPatternList()
      Dim Filename$
      List1.Clear
      For i = 0 To File1.ListCount - 1
          File1.ListIndex = i
          Filename$ = Replace(LCase(File1.Filename), ".bmp", "")
          List1.AddItem Filename$
      Next
      List1.ListIndex = 0
End Sub

Private Sub DrawFill()
  Dim txt As String
  icbFillStyle.ComboItems.Clear

  Set icbFillStyle.ImageList = imlFillStyles
  For i = 1 To 8
       Select Case i
       Case 1: txt = "Solid"
       Case 2: txt = "Transparent"
       Case 3: txt = "Horizontal Line"
       Case 4: txt = "Vertical Line"
       Case 5: txt = "Upward Diagonal"
       Case 6: txt = "Downward Diagonal"
       Case 7: txt = "Cross"
       Case 8: txt = "Diagonal Cross"
       End Select
        icbFillStyle.ComboItems.Add i, txt, txt
        icbFillStyle.ComboItems(i).Image = i
  Next i
  icbFillStyle.SelectedItem = icbFillStyle.ComboItems(1)
  Combo1.Clear
  Combo1.AddItem "Rect Horiz"
  Combo1.AddItem "Rect Vert"
  Combo1.AddItem "Tri F Diag"
  Combo1.AddItem "Tri B Diag"
  Combo1.AddItem "Rect Horiz2"
  Combo1.AddItem "Rect Vert2"
  Combo1.AddItem "Tri F Diag2"
  Combo1.AddItem "Tri B Diag2"
  Combo1.AddItem "Tri 4 Way"
  Combo1.AddItem "Tri 4 WayB"
  Combo1.AddItem "Tri 4 WayW"
  Combo1.AddItem "Tri 4 WayG"
  Combo1.ListIndex = 0

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Color1() As OLE_COLOR
    Color1 = m_Color1
End Property

Public Property Let Color1(ByVal New_Color1 As OLE_COLOR)
    m_Color1 = New_Color1
    PropertyChanged "Color1"
    PicFillColor.BackColor = m_Color1
    PictureGrd1.BackColor = m_Color1
    PicturePat1.BackColor = m_Color1
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Color2() As OLE_COLOR
    Color2 = m_Color2
End Property

Public Property Let Color2(ByVal New_Color2 As OLE_COLOR)
    m_Color2 = New_Color2
    PropertyChanged "Color2"
    
    PictureGrd2.BackColor = m_Color2
    PicturePat2.BackColor = m_Color2
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get FillStyle() As Integer
    FillStyle = m_FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As Integer)
    m_FillStyle = New_FillStyle
    PropertyChanged "FillStyle"
    If m_FillStyle <= 7 Then
       TypeFill = 1
       icbFillStyle.SelectedItem = icbFillStyle.ComboItems(m_FillStyle + 1)
    ElseIf m_FillStyle = 8 Then
       TypeFill = 2
    ElseIf m_FillStyle = 9 Then
       TypeFill = 3
    ElseIf m_FillStyle = 10 Then
       TypeFill = 4
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get NamePattern() As String
    NamePattern = m_NamePattern
End Property

Public Property Let NamePattern(ByVal New_NamePattern As String)
    Dim idFlist As Long
    
    m_NamePattern = New_NamePattern
    PropertyChanged "NamePattern"
    If TypeFill = 3 Then
      If m_NamePattern <> "" Then
         idFlist = List1.ListIndex
         For i = 0 To File1.ListCount - 1
            File1.ListIndex = i
            If LCase(File1.Filename) = LCase(m_NamePattern) Then
               List1.ListIndex = File1.ListIndex
               Exit Property
            End If
         Next
         File1.ListIndex = idFlist
      End If
     Else
       If FileExists(m_NamePattern) Then
          ReadPicturePattern m_NamePattern
       End If
     End If
    
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Color1 = m_def_Color1
    m_Color2 = m_def_Color2
    m_FillStyle = m_def_FillStyle
    m_NamePattern = m_def_NamePattern
    m_TypeFill = m_def_TypeFill
    m_TypeGradient = m_def_TypeGradient
    m_Blend = m_def_Blend
End Sub

Private Sub UserControl_Resize()
    CommandFill.Move CommandFill.Left, CommandFill.Top, CommandFill.Width, CommandFill.Height
    CommandGrd.Move CommandFill.Left, CommandFill.Top, CommandFill.Width, CommandFill.Height
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color1", m_Color1, m_def_Color1)
    Call PropBag.WriteProperty("Color2", m_Color2, m_def_Color2)
    Call PropBag.WriteProperty("FillStyle", m_FillStyle, m_def_FillStyle)
    Call PropBag.WriteProperty("NamePattern", m_NamePattern, m_def_NamePattern)
    Call PropBag.WriteProperty("TypeFill", m_TypeFill, m_def_TypeFill)
    Call PropBag.WriteProperty("BackColor", Picture1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("TypeGradient", m_TypeGradient, m_def_TypeGradient)
    Call PropBag.WriteProperty("Blend", m_Blend, m_def_Blend)
End Sub

Private Sub ReadImagedata()
    
    ReDim nColor(0)
    Image1.ScaleMode = 3
    For Y = 0 To 7
      For X = 0 To 7
          ImageData(X, Y) = Image1.POINT(X, Y)
          Findcolor = False
          For i = 1 To UBound(nColor)
              If nColor(i) = ImageData(X, Y) Then
                 Findcolor = True
                 Exit For
              End If
          Next
          If Findcolor = False Then
             ReDim Preserve nColor(UBound(nColor) + 1)
             nColor(UBound(nColor)) = ImageData(X, Y)
          End If
      Next
    Next
    If UBound(nColor) > 0 Then Color1 = nColor(1)
    If UBound(nColor) > 1 Then Color2 = nColor(2)
End Sub

Private Sub ReplaceColor(PrevColor As Long, NewColor As Long)
    For Y = 0 To 7
      For X = 0 To 7
         If ImageData(X, Y) = PrevColor Then
            ImageData(X, Y) = NewColor
            Image1.PSet (X, Y), NewColor
         End If
      Next
    Next
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicturePattern,PicturePattern,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = PicturePattern.Image
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get TypeFill() As Integer
    TypeFill = m_TypeFill
End Property

Public Property Let TypeFill(ByVal New_TypeFill As Integer)
    m_TypeFill = New_TypeFill
    PropertyChanged "TypeFill"
    Select Case m_TypeFill
    Case 1
       Option1.Value = True
    Case 2
       Option2.Value = True
    Case 3
       Option3.Value = True
    Case 4
       Option4.Value = True
    End Select
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Picture1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Picture1.BackColor() = New_BackColor
    Picture2.BackColor() = New_BackColor
    Picture3.BackColor() = New_BackColor
    Picture4.BackColor() = New_BackColor
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get TypeGradient() As Integer
    TypeGradient = m_TypeGradient
End Property

Public Property Let TypeGradient(ByVal New_TypeGradient As Integer)
    m_TypeGradient = New_TypeGradient
    PropertyChanged "TypeGradient"
    'Combo1.ListIndex = -1
    Combo1.ListIndex = m_TypeGradient
    DrawGradient
End Property

Sub mRaiseEvent()
    RaiseEvent Apply(TypeFill, FillStyle, Color1, Color2, NamePattern, TypeGradient, Blend)
End Sub

Sub DrawGradient()
    If Combo1.ListIndex = -1 Then Exit Sub
    Call GradientFillRectDC(PictureGradient.hdc, 0, 0, PictureGradient.ScaleWidth, PictureGradient.ScaleHeight, _
                            Color1, Color2, Combo1.ListIndex)
    PictureGradient.Refresh
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,255
Public Property Get Blend() As Integer
    Blend = m_Blend
End Property

Public Property Let Blend(ByVal New_Blend As Integer)
    m_Blend = New_Blend
    PropertyChanged "Blend"
    If m_Blend >= 0 And m_Blend < 255 Then
       HScroll2.Value = m_Blend
    Else
       m_Blend = 254
       HScroll2.Value = 254
    End If
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=mPicture,mPicture,-1,Image
Public Property Get PicImage() As Picture
Attribute PicImage.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set PicImage = mPicture.Image
End Property

