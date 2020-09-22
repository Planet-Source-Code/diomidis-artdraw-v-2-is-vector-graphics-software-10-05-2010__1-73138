VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmColorPicker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color "
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmColorPicker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmColorPicker.frx":030A
   MousePointer    =   1  'Arrow
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   3285
      Top             =   4380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6105
      TabIndex        =   36
      Text            =   "Choose a picture!"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6195
      TabIndex        =   35
      Top             =   3435
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.OptionButton objOption 
      Caption         =   "Image"
      Height          =   255
      Index           =   9
      Left            =   6195
      TabIndex        =   34
      Top             =   3075
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CheckBox chbPreview 
      Caption         =   "Preview"
      Enabled         =   0   'False
      Height          =   225
      Left            =   315
      TabIndex        =   28
      Top             =   4590
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picBigBox 
      Height          =   3870
      Left            =   225
      MouseIcon       =   "frmColorPicker.frx":045C
      MousePointer    =   99  'Custom
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   27
      Top             =   465
      Width           =   3885
   End
   Begin VB.PictureBox picThinBox 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00C0C0FF&
      ForeColor       =   &H00C0FFC0&
      Height          =   3840
      Left            =   4260
      Picture         =   "frmColorPicker.frx":05AE
      ScaleHeight     =   252
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   26
      Top             =   465
      Width           =   315
   End
   Begin VB.TextBox txtHexColor 
      Height          =   285
      Left            =   4995
      MaxLength       =   6
      TabIndex        =   20
      Text            =   "HexColor"
      Top             =   4230
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   6600
      TabIndex        =   19
      Text            =   "b"
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   6555
      TabIndex        =   18
      Text            =   "a"
      Top             =   2160
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   6
      Left            =   6555
      TabIndex        =   17
      Text            =   "Lab"
      Top             =   1770
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   5280
      TabIndex        =   16
      Text            =   "B"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   5280
      TabIndex        =   15
      Text            =   "G"
      Top             =   3345
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   5280
      TabIndex        =   14
      Text            =   "R"
      Top             =   2940
      Width           =   450
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   13
      Text            =   "Brightness"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   12
      Text            =   "Saturation"
      Top             =   2145
      Width           =   420
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   5265
      TabIndex        =   11
      Text            =   "Hue"
      Top             =   1755
      Width           =   435
   End
   Begin VB.OptionButton objOption 
      Caption         =   "b"
      Enabled         =   0   'False
      Height          =   255
      Index           =   8
      Left            =   6105
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton objOption 
      Caption         =   "a"
      Enabled         =   0   'False
      Height          =   255
      Index           =   7
      Left            =   6090
      TabIndex        =   9
      Top             =   2130
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton objOption 
      Caption         =   "L"
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   6105
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton objOption 
      Caption         =   "B:"
      Height          =   255
      Index           =   5
      Left            =   4785
      TabIndex        =   7
      Top             =   3735
      Width           =   495
   End
   Begin VB.OptionButton objOption 
      Caption         =   "G:"
      Height          =   255
      Index           =   4
      Left            =   4740
      TabIndex        =   6
      Top             =   3345
      Width           =   510
   End
   Begin VB.OptionButton objOption 
      Caption         =   "R:"
      Height          =   255
      Index           =   3
      Left            =   4740
      TabIndex        =   5
      Top             =   3000
      Width           =   495
   End
   Begin VB.OptionButton objOption 
      Caption         =   "B:"
      Height          =   255
      Index           =   2
      Left            =   4815
      TabIndex        =   4
      Top             =   2520
      Width           =   480
   End
   Begin VB.OptionButton objOption 
      Caption         =   "S:"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   3
      Top             =   2070
      Width           =   480
   End
   Begin VB.OptionButton objOption 
      Caption         =   "H:"
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   2
      Top             =   1785
      Width           =   465
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6225
      TabIndex        =   1
      Top             =   600
      Width           =   885
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      Height          =   345
      Left            =   6225
      TabIndex        =   0
      Top             =   150
      Width           =   885
   End
   Begin VB.Image imgRetro 
      Height          =   420
      Left            =   1875
      Picture         =   "frmColorPicker.frx":41F2
      Stretch         =   -1  'True
      Top             =   4410
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgWinter 
      Height          =   495
      Left            =   1320
      Picture         =   "frmColorPicker.frx":16621
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblPicPath 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   3930
      TabIndex        =   39
      Top             =   4590
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Line linTriang2Falling 
      X1              =   318
      X2              =   327
      Y1              =   184
      Y2              =   189
   End
   Begin VB.Line linTriang2Rising 
      X1              =   325
      X2              =   334
      Y1              =   195
      Y2              =   189
   End
   Begin VB.Line linTriang2Vert 
      X1              =   318
      X2              =   315
      Y1              =   185
      Y2              =   200
   End
   Begin VB.Label lblThinContainer 
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000007&
      Height          =   3870
      Left            =   4200
      TabIndex        =   38
      Top             =   465
      Width           =   525
   End
   Begin VB.Line linTriang1Falling 
      X1              =   277
      X2              =   282
      Y1              =   251
      Y2              =   256
   End
   Begin VB.Line linTriang1Rising 
      X1              =   277
      X2              =   282
      Y1              =   261
      Y2              =   256
   End
   Begin VB.Line linTriang1Vert 
      X1              =   277
      X2              =   277
      Y1              =   251
      Y2              =   261
   End
   Begin VB.Label Label4 
      Caption         =   "Recent images"
      Height          =   300
      Left            =   6030
      TabIndex        =   37
      Top             =   3930
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblComplementaryColor 
      BackColor       =   &H80000017&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C.C."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   5745
      TabIndex        =   33
      ToolTipText     =   "Complementary Color (adds 180° to Hue Angle)."
      Top             =   645
      Width           =   435
   End
   Begin VB.Label lblSuffix 
      Caption         =   "%"
      Height          =   270
      Index           =   2
      Left            =   5790
      TabIndex        =   31
      Top             =   2550
      Width           =   210
   End
   Begin VB.Label lblSuffix 
      Caption         =   "%"
      Height          =   270
      Index           =   1
      Left            =   5775
      TabIndex        =   30
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label lblSuffix 
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   5775
      TabIndex        =   29
      Top             =   1755
      Width           =   210
   End
   Begin VB.Image imgMarker 
      Height          =   165
      Left            =   2535
      Picture         =   "frmColorPicker.frx":256D1
      Top             =   4560
      Width           =   165
   End
   Begin VB.Label lblOldColor 
      BackColor       =   &H00FFFF80&
      Height          =   495
      Left            =   4800
      TabIndex        =   25
      Top             =   1005
      Width           =   900
   End
   Begin VB.Label lblNewColor 
      Appearance      =   0  'Flat
      BackColor       =   &H0099CCDD&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4800
      TabIndex        =   24
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   15
      Left            =   5475
      TabIndex        =   23
      Top             =   1845
      Width           =   15
   End
   Begin VB.Label Label2 
      Caption         =   "Select color:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4785
      TabIndex        =   21
      Top             =   4245
      Width           =   195
   End
   Begin VB.Label lblContainer 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1050
      Left            =   4755
      TabIndex        =   32
      Top             =   465
      Width           =   690
   End
End
Attribute VB_Name = "frmColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Canceled As Boolean
Dim blnDrag As Boolean, intSystemColorAngleMax1530 As Integer, bteSaturationMax255 As Byte, bteBrightnessMax255 As Byte
Dim mSngRValue As Single, mSngGValue As Single, mSngBValue As Single
Dim blnNotFirstTimeMarker As Boolean, mBteMarkerOldX As Integer, mBteMarkerOldY As Integer
Dim arLongMarkerColorStore(11, 11) As Long, arsPicPath() As String
Dim mBlnRecentThinBoxPress As Boolean, mBlnBigBoxReady As Boolean
'Welcome to use, improve and share this utility. It gives you more control than the standard vb-colorpicker.
'Anna-Carin who created this program gives it away for free.
'A SMALL BUG TO FIX IS THAT THE NUDGE FUNCTION OF THE ARROW MARKERS LOSES FOCUS.

'API TO PAINT PIXELS IN picBoxes.
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Byte 'Painting by API is good and fast.

'FINDS THE REAL PATH FOR MyDocuments
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As String) As Long

'Shelling html
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type HSL 'IS USED FOR THE HSL FUNCTION FROM THE WEBSITE VBspeed.
    Hue As Integer 'FROM 0 To 360.
    Saturation As Byte
    Luminance As Byte
End Type


Function Shell(Program As String, Optional ShowCmd As Long = vbNormalNoFocus, Optional ByVal WorkDir As Variant) As Long

Dim FirstSpace As Integer, Slash As Integer

If Left(Program, 1) = """" Then
    FirstSpace = InStr(2, Program, """")
    If FirstSpace <> 0 Then
        Program = Mid(Program, 2, FirstSpace - 2) & Mid(Program, FirstSpace + 1)
        FirstSpace = FirstSpace - 1
    End If
Else
    FirstSpace = InStr(Program, " ")
End If

If FirstSpace = 0 Then FirstSpace = Len(Program) + 1

If IsMissing(WorkDir) Then
    For Slash = FirstSpace - 1 To 1 Step -1
        If Mid(Program, Slash, 1) = "\" Then Exit For
    Next

If Slash = 0 Then
    WorkDir = CurDir
ElseIf Slash = 1 Or Mid(Program, Slash - 1, 1) = ":" Then
    WorkDir = Left(Program, Slash)
Else
    WorkDir = Left(Program, Slash - 1)
End If
End If

Shell = ShellExecute(0, vbNullString, _
Left(Program, FirstSpace - 1), LTrim(Mid(Program, FirstSpace)), _
WorkDir, ShowCmd)
If Shell < 32 Then VBA.Shell Program, ShowCmd 'To raise Error

End Function

Private Sub Command1_Click()
      Canceled = False
      Hide
End Sub

Public Sub Form_Load()
Dim udtAngelSaturationBrightness As HSL, bteValdRadioKnapp As Byte
'ReDim Preserve arsPicPath(1) 'arsPicPath NEEDS A FIRST INITIALISATION TO ENABLE THE USE OF Ubound LATER.

Dim Ctr As Byte, bteExtraWidth As Byte, bteExtraHeight As Byte
'STYLING THE FORM. STRANGELY FAILED TO SWITCH THE SCALEMODE TO PIXLES. 1 pixel=20 twips.

mBlnRecentThinBoxPress = True 'TO GET RID OF GREY SQUARES IN THE PICTURE.
'frmColorPicker.ScaleMode = vbPixels 'RESEMBLING PIXELS.
frmColorPicker.Width = 7380
frmColorPicker.Height = 5280

chbPreview.Move 13, 299, 103, 15 'Left,Top,Width,Height.
chbPreview.Visible = False 'HIDES IT IN THIS BETAVERSION CAUSE IT DOESN'T YET HAVE A PURPOSE.
imgMarker.Visible = False 'HIDING TO IMPROVE THE LOOKS.

For Ctr = 0 To 2
    Text1(Ctr).Move 351, 117 + Ctr * 25, 30, 21
    objOption(Ctr).Move 320, 120 + Ctr * 25, 33, 17
Next Ctr

For Ctr = 3 To 5
    Text1(Ctr).Move 350, 196 + (Ctr - 3) * 26, 30, 21
    objOption(Ctr).Move 320, 198 + (Ctr - 3) * 26, 33, 17
Next Ctr

For Ctr = 6 To 8
    Text1(Ctr).Move 437, 117 + (Ctr - 6) * 26, 34, 21
    objOption(Ctr).Move 407, 120 + (Ctr - 6) * 25, 25, 17
Next Ctr
    
    txtHexColor.Move 336, 281, 56, 19

    Label1.Move 319, 283, 13, 14 'tecknet #
        
    lblNewColor.Move 322, 33, 58, 33
    lblOldColor.Move 322, 66, 58, 33
    lblContainer.Move 321, 32, 60, 68
    lblOldColor.BackColor = lblNewColor.BackColor 'STARTS AT THE SAME COLOR.
    
    picThinBox.Picture = Nothing 'LOADS WITH NOTHING TO GET GET THE JPG-IMAGE OUT OF SIGHT.
    picThinBox.ScaleMode = vbPixels
'CALCULATES THE FRAME WIDTH BEFORE STYLING.
bteExtraWidth = picThinBox.Width - picThinBox.ScaleWidth 'OUTER MEASURE MINUS ACTUAL INNER MEASURE = FRAMEWIDTH.
bteExtraHeight = picThinBox.Height - picThinBox.ScaleHeight 'OUTER MEASURE MINUS ACTUAL INNER MEASURE = FRAMEWIDTH.
picThinBox.Move 284, 31, 19 + bteExtraWidth, 256 + bteExtraHeight 'FRAMES ARE 4 UNITS BROAD. CURIOSITY FACT IS THAT TEH FRAMES OF ALL VBCONTROLS EXCEPT FOR forms ARE MEASURED FROM THE FRAME CENTER, SO YOU ACTUALLY GET HALF THE WIDTH, BUT IT WORKS SINCE VB USE THE SAME LOGIC ALL THE WAY.

lblThinContainer.BackStyle = 0 'Transparent
lblThinContainer.Left = 284 - 10: lblThinContainer.Top = picThinBox.Top: lblThinContainer.Width = picThinBox.Width + 20: lblThinContainer.Height = picThinBox.Height
    
linTriang1Vert.X1 = 277: linTriang1Vert.X2 = 277: linTriang1Vert.Y1 = 251: linTriang1Vert.Y2 = 261
linTriang1Rising.X1 = 277: linTriang1Rising.X2 = 283: linTriang1Rising.Y1 = 261: linTriang1Rising.Y2 = 256
linTriang1Falling.X1 = 277: linTriang1Falling.X2 = 283: linTriang1Falling.Y1 = 251: linTriang1Falling.Y2 = 256

linTriang2Vert.X1 = 314: linTriang2Vert.X2 = 314: linTriang2Vert.Y1 = 251: linTriang2Vert.Y2 = 261
linTriang2Rising.X1 = 309: linTriang2Rising.X2 = 314: linTriang2Rising.Y2 = 261: linTriang2Rising.Y1 = 256
linTriang2Falling.X1 = 309: linTriang2Falling.X2 = 314: linTriang2Falling.Y2 = 251: linTriang2Falling.Y1 = 256

    
    picBigBox.Width = 256 + 4 '256 INCREASING BY 4 SINCE VB PROBABLY CHEATS THE SAME WAY AS IT DID IN picThinBox.
    picBigBox.Height = 256 + 4
    picBigBox.ScaleWidth = 256
    picBigBox.ScaleHeight = 256
    picBigBox.Left = 13
    picBigBox.Top = 31

lblPicPath.Left = 13: lblPicPath.Width = 460

'objOption(0) = True 'STATES Hue AS DEFAULT. ***ATT!!!!!  THIS BOOTS THE CLICK ROUTINE TO DECORATE ThinBox AND BigBox.
Call SplitlblNewColorToRGBboxes 'ALSO THE SYSTEM CONSTANTS OF RGB GETS UPDATED.
udtAngelSaturationBrightness = RGBToHSL201(lblNewColor.BackColor, True) 'TRUE MEANS THAT HSL IS UPDATING BOTH THE textboxes AND THE systemConstants.

'Call ExecuteIniFile(bteValdRadioKnapp) 'Chooses the latest mode of optRadioButton.
objOption(bteValdRadioKnapp) = True

End Sub

' Display the form. Return True if the user cancels.
Public Function ShowForm(ByRef mColor As Long) As Boolean
Dim t As Integer, IndexS As Integer, Idx As Integer
    Dim mr As Long, mg As Long, mb As Long
    ' Assume we will cancel.
    Canceled = False
    lblNewColor.BackColor = mColor
    SplitRGB mColor, mr, mg, mb
    Text1(3).Text = Str(mr)
    Text1(4).Text = Str(mg)
    Text1(5).Text = Str(mb)
    Text1_LostFocus 3
    Text1_LostFocus 4
    Text1_LostFocus 5
    Screen.MousePointer = 0
    
    ' Display the form.
    Show vbModal

    ShowForm = Canceled
    mColor = lblNewColor.BackColor
    Unload Me
End Function

Private Sub objOption_Click(Index As Integer) 'Choosing modus.
Dim Ctr As Integer
If Index <> 9 And txtHexColor.Left = 286 Then ' Restore HexBox & Combo1.
lblPicPath.Visible = False
cmdBrowse.Enabled = False: Combo1.Enabled = False
'MsgBox "Move HexBox"
For Ctr = 286 To 336
    txtHexColor.Move Ctr, 281, 56, 20
    Combo1.Move Ctr + 70, 281, 70 + 336 - Ctr 'Height-property in ComboBoxes is readonly.
Next Ctr
DoEvents 'Problems with visual jam.
End If


If Index = 0 Then 'MsgBox "Hue"
    'picThinBox.Visible = True
    Call PaintThinBox(0)
    mBteMarkerOldX = bteSaturationMax255: mBteMarkerOldY = 255 - bteBrightnessMax255
    If mBlnBigBoxReady = True Then Call picBigBox_Colorize 'NO, MAKE THIS EASIER - REDRAW ONLY IF setup HAS FINISHED.
    Call picBigBox_Colorize
End If

If Index = 1 Then
    'picThinBox.Visible = True
    Call PaintThinBox(1)
    mBteMarkerOldX = intSystemColorAngleMax1530 / 6: mBteMarkerOldY = 255 - bteBrightnessMax255
    Call picBigBox_Colorize 'Speciell design.
    End If

If Index = 2 Then ' "Brightness"
    Call PaintThinBox(2)
    mBteMarkerOldX = intSystemColorAngleMax1530 / 6: mBteMarkerOldY = 255 - bteSaturationMax255
    Call picBigBox_Colorize 'Speciell design.
End If

If Index = 3 Then ' "R"
    Call opt3RedPaintPicThinBox(ByVal Text1(4), Text1(5))
    mBteMarkerOldX = Text1(5): mBteMarkerOldY = 255 - Text1(4)
    Call picBigBox_Colorize 'Speciell design.
    
End If
If Index = 4 Then ' "G"
    Call opt4GreenPaintPicThinBox(ByVal Text1(3), Text1(5))
    mBteMarkerOldX = Text1(5): mBteMarkerOldY = 255 - Text1(3)
    Call picBigBox_Colorize 'Speciell design.

End If
If Index = 5 Then ' "B"
    Call opt5BluePaintPicThinBox(ByVal Text1(3), Text1(4))
    mBteMarkerOldX = Text1(3): mBteMarkerOldY = 255 - Text1(4)
    Call picBigBox_Colorize 'Speciell design.

End If

If Index = 9 Then ' "PictureBrowse"
    blnNotFirstTimeMarker = False 'LOSES AN ALIEN SUGAR CUBE IN BigBox.
    picThinBox.Visible = False 'BackColor = HSLToRGB(ByVal intSystemColorAngleMax1530, ByVal bteSaturationMax255, ByVal 255, False) 'Gets a lighter shade of the active color. 'Sets the whole square for easy fading.
    lblPicPath.Visible = True: cmdBrowse.Enabled = True: Combo1.Enabled = True
    Call MoveHexBox
End If

Call imgArrowsModeDepending 'MOVING imgArrows

picThinBox.Refresh
End Sub

Private Sub Combo1_Click() 'THE NUMERIC BASE OF THE LIST IS ZERO.
Dim Answer As Integer, sFile As String
Static bteListIndex As Byte 'IS USED TO REMARK IN THE BOX.
On Error GoTo ErrorHandler 'Error 53= File doesnt exist.
bteListIndex = Combo1.ListIndex 'IS USED TO REMARK IN THE BOX.
If Combo1.ListIndex = 0 Then picBigBox = imgWinter: lblPicPath = "Native picture."
If Combo1.ListIndex = 1 Then picBigBox = imgRetro: lblPicPath = "Native picture."
If Combo1.ListIndex > 1 Then
    picBigBox.Picture = LoadPicture(arsPicPath(Combo1.ListIndex - 1))
    lblPicPath = arsPicPath(Combo1.ListIndex - 1)
End If
mBlnBigBoxReady = True 'THERE ARE COLORS IN bigbox.
Exit Sub

ErrorHandler: 'In case file is missing
If Err.Number = 76 Then '76 IS THE CORRECT error number. Number 53 IS USED FOR SPECIAL FILEHANDLING SITUATION.
    Answer = MsgBox("Broken link! Do you want to manually search for the picture  " & Combo1.Text & "  on your drive?", vbQuestion + vbYesNo, "Broken Link Action")
    If Answer = vbYes Then
        Call RepairLink(sFile)
        If sFile = "Cancel" Then Resume Next 'CONDITIONAL JUMP TO ABOVE.
        Combo1.ListIndex = bteListIndex 'MAKES A MARKING IN THE BOX TO SHOW THE IMAGE. IT'S THE SAME THING AS CLICKING IN TEH BOX AND THEREFOR THE PROGRAM LEAPS TO THE START OF THIS ROUTINE, BUT I THINK IT IS OK THEN.
    End If
End If

'THE USER DID NOT WANT TO BROWSE FOR THE IMAGE.

Answer = MsgBox("Do you want the computer to automatically scan the whole image list and remove all broken links? (Only the image list itself will be affected - your hard drive will be left alone!)", vbQuestion + vbYesNo, "Broken Link Action")
If Answer = vbYes Then Call RemoveBrokenLinks: Exit Sub
End Sub
Public Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Ctr As Byte, Answer As Integer, bteListIndex As Byte

If KeyCode <> vbKeyDelete Then Exit Sub
bteListIndex = Combo1.ListIndex 'THE HIGHLIGHTED POST IN Combo1.
If bteListIndex < 2 Then MsgBox "This picture is native within the programme kan can not be deleted": Exit Sub
Answer = MsgBox("Are you sure you want remove  " & Combo1.Text & "  from the list?", vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then Exit Sub

Call ContractArrays(bteListIndex)

picBigBox = Nothing

End Sub

Public Sub cmdBrowse_Click()
Dim sFile As String, strPicName As String
Call OpenDialog(sFile)
lblPicPath = sFile 'Shows the path on tje long lable.

ReDim Preserve arsPicPath(UBound(arsPicPath) + 1) 'THE MATRIX GROWS WITH THE SIZE OF OEN ENTITY.
arsPicPath(UBound(arsPicPath)) = sFile 'STORING THE WHOLE PATH.

'EXTRACTING THE NAME OF THE IMAGEFILE + extension from the path.
strPicName = Mid(sFile, InStrRev(sFile, "\") + 1) 'Extraherar filnamnet ur pathen, dock med.wav.
If strPicName = "-" Then 'Evalueating.
    strPicName = "": sFile = "": MsgBox "Unvalid filename. The filename is - and that's no good if you want to keep track of a namelist. I cancel!": Exit Sub
End If
picBigBox.Picture = LoadPicture(sFile)
Combo1.AddItem strPicName 'ADDING THE IMAGE NAME TO THE COMBOBOX .

'ApplPath
End Sub

Private Sub cmdCancel_Click()
    Canceled = True
    Unload Me
End Sub

Public Sub picBigBox_Colorize()
Dim blnFadeToGrey As Boolean, R As Single, G As Single, B As Single

    If objOption(9) Then Exit Sub 'IN CASE Option(9) THE bigbox SHALL BE LEFT ALONE.

    picBigBox.Visible = False 'HIDES THE SLOW DRAWIING PROCEDURE.
    picBigBox.AutoRedraw = True 'ELSE YOU WONT SEE ANYTHING.

    'Set objAnyPictureBox = picBigBox 'Ritar om canvas.
    '*****     ********     **********     **********
    If objOption(0) Then 'IN CASE Option(0) WE SHALL FETCH a fully saturated version of color AND MAKE A 3-D FADE.'
        Call Bigbox3D 'NEW VERSION
    End If

If objOption(1) Then
    Call RainBowBigbox(vbFalse, vbTrue) 'FadeToGrey=False & FadeToBlack=True
End If
If objOption(2) Then
    Call RainBowBigbox(vbTrue, vbFalse) 'FadeToGrey= true & FadeToBlack=false
End If

If objOption(3) Then
    Call opt3RedPaintPicBigBox
End If
If objOption(4) Then
    Call opt4GreenPaintPicBigBox
End If
If objOption(5) Then
    Call opt5BluePaintPicBigBox
End If

picBigBox.Visible = True 'SHOWS THE PICBOX AFTER THE SLOW DECORATION.
picThinBox.Visible = True 'IS NEEDED TO SHOW IN CASE THE FORMER MODE WAS POSTCARDVIEW WHICH THUS HIDES ThinBox.

If blnNotFirstTimeMarker = True Then 'IN CASE THERE IS A marker-coordinate...
    Call SampleMarkerBackground   'SAVES THE BACKGROUND OF MARKER IF THERE IS ANY.
    Call PaintMarker(mBteMarkerOldX, mBteMarkerOldY) 'REPAINT THE MARKER (if there is any).
    lblNewColor.BackColor = picBigBox.POINT(mBteMarkerOldX, mBteMarkerOldY)
End If
If mBlnBigBoxReady = False Then 'PLACES A MARKER AT CORRECT LOCATION AT THE SETUP STAGE.
    blnNotFirstTimeMarker = True 'PASSWORD.
    'MODE DEPENDING NEW MARKER POSITION.
    If objOption(0) Then mBteMarkerOldX = bteSaturationMax255:   mBteMarkerOldY = 255 - bteBrightnessMax255 'Transmitting logical values.
    If objOption(1) Then mBteMarkerOldX = intSystemColorAngleMax1530 / 6: mBteMarkerOldY = 255 - bteBrightnessMax255 'Transmitting logical values.
    If objOption(2) Then mBteMarkerOldX = intSystemColorAngleMax1530 / 6: mBteMarkerOldY = 255 - bteSaturationMax255 'Transmitting logical values.

    Call SampleMarkerBackground   'SAVES THE BACKGROUND OF MARKER IF THERE IS ANY.
    Call PaintMarker(mBteMarkerOldX, mBteMarkerOldY) 'REPAINT THE MARKER (if there is any).
    mBlnBigBoxReady = True 'NOW AT LEAST THE FIRST SPONTANEOUS REDRAW HAS FINISHED.
End If
End Sub

Private Sub picBigBox_MouseMove(Knapp As Integer, Shift As Integer, X As Single, Y As Single)
'PROBLEM: GIF-IMAGES ETC WONT REACT WHEN I SAVE THE OLD IMAGE AS A MATRIX. ON THE OTHER HAND I CAN PAINT OVER GIFS.
Dim lngColor As Long, udtAngelSaturationBrightness As HSL

If blnDrag = False Then Exit Sub 'Baile if mousebutton is not held down.

If X > 255 Then X = 255 'LIMITER.
If X < 0 Then X = 0
If Y > 255 Then Y = 255
If Y < 0 Then Y = 0

'PASTE THE MARKER ON THE LOCATION OF X,Y.*******
'HIDE THE MARKER FOR CONVENIENS. LET THE MARKER FOLLOW IF THE MOUSEBUTTON IS PRESSED. FIRST ERASE THE OLD MARKER.
If objOption(0) Then lngColor = HSLToRGB(intSystemColorAngleMax1530, ByVal X, ByVal 255 - Y, True) 'CONVERT AND UPDATE TEXTBOXES.
If objOption(1) Then lngColor = HSLToRGB(ByVal X * 6, ByVal bteSaturationMax255, ByVal 255 - Y, True): Call PaintThinBox(1) 'CONVERT AND UPDATE TEXTBOXES.
If objOption(2) Then lngColor = HSLToRGB(ByVal X * 6, ByVal 255 - Y, ByVal bteBrightnessMax255, True): Call PaintThinBox(2) 'CONVERT AND UPDATE TEXTBOXES.

If objOption(3) Then Call BigBoxOpt3Reaction(ByVal X, Y) 'CONVERT AND UPDATE TEXTBOXES.
If objOption(4) Then Call BigBoxOpt4Reaction(ByVal X, Y) 'CONVERT AND UPDATE TEXTBOXES.
If objOption(5) Then Call BigBoxOpt5Reaction(ByVal X, Y) 'CONVERT AND UPDATE TEXTBOXES.

If objOption(9) Then lblNewColor.BackColor = picBigBox.POINT(X, Y): Call SplitlblNewColorToRGBboxes: udtAngelSaturationBrightness = RGBToHSL201(lblNewColor.BackColor, True) 'True means that HSL uppdates both the textboxes and the system constants.

mBlnRecentThinBoxPress = False
End Sub
Private Sub picBigBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngColor As Long

If mBlnBigBoxReady = False Then MsgBox "mBlnBigBoxReady = False i BigBox MouseDown! There are no colors to show in bigbox": Exit Sub 'Bail if no color in bigbox.
blnDrag = True
'HIDING THE MARKER NOT TO RISK OF GETTING JAM IN MY PROBE.
If blnNotFirstTimeMarker = True Then
    Call EraseMarker
End If

If objOption(0) Then
   lngColor = HSLToRGB(intSystemColorAngleMax1530, ByVal X, ByVal 255 - Y, True) 'CONVERT AND UPDATE TEXTBOXES.
End If
If objOption(1) Then
    lngColor = HSLToRGB(ByVal X * 6, ByVal bteSaturationMax255, ByVal 255 - Y, True) 'CONVERT AND UPDATE TEXTBOXES.
    Call FadeThinBoxToGrey 'REPAINT ThinBox - FADE SATURATED COLORS; THE SYSTEM CONSTANTS ARE ALREADY UPDATED.
    picThinBox.Refresh
End If

If objOption(2) Then
    picThinBox.BackColor = HSLToRGB(ByVal X * 6, ByVal 255 - Y, 255, False) 'SETTING THE BRIGHT COLOR THAT IS TO BE FADED. CONVERTING AND UPDATING TEXTBOXES.
    lngColor = HSLToRGB(ByVal X * 6, ByVal 255 - Y, ByVal bteBrightnessMax255, True)  'UPDATING THE REAL, NONSATURATED SYSTEM CONSTANTS AND lblNewColor.
    Call FadeThinBoxToBlack 'REPAINTING ThinBox - FADE SATURATED COLORS ; THE SYSTEM CONSTANTS ARE ALREADY UPDATED.
    picThinBox.Refresh
End If

If objOption(3) Then
    Call BigBoxOpt3Reaction(ByVal X, Y)
End If
If objOption(4) Then
    Call BigBoxOpt4Reaction(ByVal X, Y)
End If
If objOption(5) Then
    Call BigBoxOpt5Reaction(ByVal X, Y)
End If
If objOption(9) Then Call picBigBox_MouseMove(Button, Shift, X, Y)
'THESE FILTER OPTIONS ARE SOEWHAT UNPREDICTABLE SO I FEEL MY WAY.
'vbSrcInvert FOLLOWED BY vbDstInvert APPARENTLY GIVES A TRANSPARENT PICTURE.
'?OLD CODE? IN CASE OF THE CURSOR COLLIDING WITH THE MARKER, THE MARKER HAS TO ENTIRELY ERASED AND ENTIRELY REPAINTED.
    mBteMarkerOldX = X 'Already changing here in order to get the correct position in SampleMarkerBackground .
    mBteMarkerOldY = Y 'Will be used by erasemarker.
Call SampleMarkerBackground 'Saving the new backround behind marker now when there's no Cursor in the way.

blnNotFirstTimeMarker = True

End Sub
Private Sub picBigBox_MouseUp(Knapp As Integer, Shift As Integer, X As Single, Y As Single)

If mBlnBigBoxReady = False Then MsgBox "mBlnBigBoxReady = False!": Exit Sub 'Baile if no color in bigbox.

blnDrag = False
If X > 255 Then X = 255 'LIMITER
If X < 0 Then X = 0
If Y > 255 Then Y = 255
If Y < 0 Then Y = 0

mBteMarkerOldX = X
mBteMarkerOldY = Y
Call SampleMarkerBackground
Call PaintMarker(X, Y) 'PAINT MARKER ON ITS NEW LOCATION.
End Sub
Public Sub EraseMarker()
'If blnSetup
Dim CtrY As Byte, CtrX As Byte
    For CtrY = 0 To 10
    For CtrX = 0 To 10
    picBigBox.PSet (mBteMarkerOldX - 5 + CtrX, mBteMarkerOldY - 5 + CtrY), arLongMarkerColorStore(CtrX, CtrY)
    Next CtrX
    Next CtrY
End Sub

Private Sub picThinBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' set flag to start drawing
mBlnRecentThinBoxPress = True
blnDrag = True: Call picThinBox_MouseMove(Button, Shift, X, Y) 'REUSING THE UPDATE ROUTINES.

End Sub
Private Sub lblThinContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sngScaleConst As Single
sngScaleConst = Screen.TwipsPerPixelY 'GIVING ME THE ACTUAL SIZE OF THE PIXELS OF THE SCREEN, HERE = 15.

mBlnRecentThinBoxPress = True
Y = Y / sngScaleConst 'CONVERTING FROM THE UNIT TWIP TO PIXELS. ATT! PROBLEM! SHOULD BE /20 BUT IS 15.
blnDrag = True: Call picThinBox_MouseMove(Button, Shift, X, Y) 'REUSING THE UPDATE ROUTINES.

End Sub
Private Sub lblThinContainer_MouseMove(Knapp As Integer, Shift As Integer, X As Single, Y As Single)
Y = Y / 15 'CONVERTING FROM THE UNIT TWIP TO PIXELS. ATT! PROBLEM! SHOULD BE /20 BUT IS 15.
Call picThinBox_MouseMove(Knapp, Shift, X, Y)
End Sub

Private Sub picThinBox_MouseMove(Knapp As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngColor As Long, udtAngelSaturationBrightness As HSL

If blnDrag = False Then Exit Sub
'If Text1(1) = "Saturation" Then Text1(1) = 100 'The program har been started recently.
If Y < 0 Then Y = 0 'LIMITER
If Y > 255 Then Y = 255
'imgArrows.Top = Y + 28 'Animering
Call TriangelMove(Y) 'ANIMATION

If objOption(0) Then lngColor = HSLToRGB((255 - Y) * 6, ByVal bteSaturationMax255, ByVal bteBrightnessMax255, True): Exit Sub 'Convert and update textboxes.
If objOption(1) Then lngColor = HSLToRGB(ByVal intSystemColorAngleMax1530, 255 - Y, ByVal bteBrightnessMax255, True): Exit Sub 'Convert and update textboxes.
If objOption(2) Then lngColor = HSLToRGB(ByVal intSystemColorAngleMax1530, ByVal bteSaturationMax255, 255 - Y, True) 'Convert and update textboxes.
If objOption(3) Then
    Text1(3) = 255 - Y: udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True) 'Convert and update textboxes.
    lblNewColor.BackColor = RGB(Text1(3), Text1(4), Text1(5))
End If
If objOption(4) Then
    Text1(4) = 255 - Y: udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True) 'Convert and update textboxes.
    lblNewColor.BackColor = RGB(Text1(3), Text1(4), Text1(5))
End If
If objOption(5) Then
    Text1(5) = 255 - Y: udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True) 'Convert and update textboxes.
    lblNewColor.BackColor = RGB(Text1(3), Text1(4), Text1(5))
End If

End Sub
Private Sub picThinBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' set flag to start drawing
blnDrag = False
Call picBigBox_Colorize
End Sub
Private Sub lblThinContainer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' set flag to start drawing
Y = Y / 20 'CONVERTING FROM THE UNIT TWIP TO PIXELS.
blnDrag = False
Call picBigBox_Colorize
End Sub

Public Sub FadeThinBoxToGrey()
Dim sng255saturation As Single, sngLokalBrightness As Single, X As Byte, Y As Integer ', YCtr As Integer

sng255saturation = 255: sngLokalBrightness = bteBrightnessMax255
    
For X = 0 To 19
    Y = 0 'Sets YCtr for making a new countdown.
    Do 'Interesting if there would raise an error, thus a leap directly to EndSub.
    SetPixelV picThinBox.hDC, X, Y, HSLToRGB(ByVal intSystemColorAngleMax1530, ByVal Round(sng255saturation - sng255saturation * Y / 255), ByVal sngLokalBrightness, False)
    Y = Y + 1
    Loop While Y < 256 'Because Y gets to big when the loop has finished.
Next X

End Sub

Public Sub Bigbox3D()
Dim sngLokalSaturation As Single, sngLokalBrightness As Single, YRADNOLL As Integer
Dim sngR256delToBlack As Single, sngG256delToBlack As Single, sngB256delToBlack As Single
Dim R As Single, G As Single, B As Single, lColor As Long, Y As Integer, X As Integer

sngLokalSaturation = 255: sngLokalBrightness = 255 'There is a need for intense start color.
'If R > G Then lSuperior = R Else lSuperior = G 'Det skulle gå att halvera denna rutin medelst en superior, men koden blir då svårare att fatta.
'If B > lSuperior Then lSuperior = B
'********* Firstly a single fade from saturated to grey on the uppermost row.
    For X = 0 To 255
    SetPixelV picBigBox.hDC, X, YRADNOLL, HSLToRGB(ByVal intSystemColorAngleMax1530, ByVal Round(sngLokalSaturation * X / 255), ByVal sngLokalBrightness, False)
    Next X 'Resets Y for a new row.

'********* Here will be an FADE TO BLACK for all columns ********

For X = 255 To 0 Step -1
lColor = picBigBox.POINT(X, 0) 'Reading the uppermost pixel which is to be faded.
R = lColor And &HFF
G = (lColor And &HFF00&) \ &H100&
B = (lColor And &HFF0000) \ &H10000
sngR256delToBlack = R / 255  'The fraction blocks which lead down to black.
sngG256delToBlack = G / 255
sngB256delToBlack = B / 255
For Y = 0 To 255 'Interesting if there would raise an error, thus a leap back to EndSub.
    SetPixelV picBigBox.hDC, X, Y, RGB(R, G, B) 'Painting with API.
    R = R - sngR256delToBlack 'Darkening the shade one of a 256:th.
    G = G - sngG256delToBlack
    B = B - sngB256delToBlack
Next Y
Y = Y - 1 'Because that Y gets too big when the loop is completed.
Next X

End Sub

Public Sub FadeThinBoxToBlack()
Dim sngR256delToBlack As Single, sngG256delToBlack As Single, sngB256delToBlack As Single
Dim R As Single, G As Single, B As Single, lColor As Long, X As Byte, Y As Integer

For X = 0 To 19
lColor = picThinBox.POINT(X, 0) 'Reads the uppermost pixel MAX LIGHT which is to be faded.
R = lColor And &HFF
G = (lColor And &HFF00&) \ &H100&
B = (lColor And &HFF0000) \ &H10000
sngR256delToBlack = R / 255  'Fractions which leads down to black.
sngG256delToBlack = G / 255
sngB256delToBlack = B / 255
For Y = 0 To 255 'Interesting if the is an error, thus a jump directly to EndSub.
    SetPixelV picThinBox.hDC, X, Y, RGB(R, G, B) 'Painting with API.
    R = R - sngR256delToBlack 'Darkening the shade of one 256th.
    G = G - sngG256delToBlack
    B = B - sngB256delToBlack
Next Y
Y = Y - 1 'Because Y gets too big when loop is complete.
Next X

End Sub
Public Sub RainBowBigbox(blnFadeToGrey, blnFadeToBlack) 'Is used by both radiobutton 1 & 2.
Dim Ctr As Byte, blnUpdateTextBoxes As Boolean, bteK4243 As Byte
Dim Saturation As Single, Luminance As Single
Static intNODE As Integer, YCtr As Integer, XCtr As Integer, intRainbowAngle As Integer
    'There is no risk for getting dull shades since I use the native principal by adding/subtracting values against at constant FF-component.
    'The algoritm gives med decimal values which increases the importance for mathematical models for choosing color, not pic.point.
    intRainbowAngle = 0 'Protects the systemcolorangel

    If blnFadeToGrey = vbTrue And blnFadeToBlack = vbFalse Then
        Saturation = 255
        Luminance = bteBrightnessMax255 'Starting value fully saturated. Brightness is to be the same for the whole of bigbox.
    Else
        Saturation = bteSaturationMax255
        Luminance = 255 'Fading from fully bright.
    End If

    bteK4243 = 42 'Has to alternate between 42 and 43 pixels per colorfield to make even at 256 pixels.

    XCtr = 0 'To255
    For YCtr = 0 To 255
        Do 'X loopen 0 To 255.
            '1 Red in in direction towards yellow. Green is counting up.
            For Ctr = 1 To bteK4243  'Has to alternate between 42 and 43 pixels per colorfield to make even at 256 pixels.
                If blnFadeToBlack Then Luminance = 255 - YCtr 'Round(bteBrightnessMax255 - (bteBrightnessMax255 / 255 * YCtr))
                If blnFadeToGrey Then Saturation = 255 - YCtr 'Round(bteSaturationMax255 - (bteSaturationMax255 / 255 * YCtr))
                intRainbowAngle = intNODE + ((254 * (Ctr - 1)) / (bteK4243 - 1)) 'Wonderful solution: this logic about going from zero to the full value (here 254) I have been seeking for a long time.
                SetPixelV picBigBox.hDC, XCtr, YCtr, HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, False)
                XCtr = XCtr + 1
            Next Ctr '
            If bteK4243 = 43 Then bteK4243 = 42 Else bteK4243 = 43
            intNODE = intNODE + 255 'Bistabile switch.
        Loop While XCtr < 255
        intRainbowAngle = 0 'Painting the last fully red which lies outside the logic.
        picBigBox.PSet (XCtr, YCtr), HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, blnUpdateTextBoxes)
        intNODE = 0: XCtr = 0
    Next YCtr
End Sub

Public Sub RainBowThinBox() 'By swapping the XY-vvalues at the call you can paint either horisontal or vertical.
Dim Ctr As Byte, blnUpdateTextBoxes As Boolean, bteK4243 As Byte
Dim blnHorizontal As Boolean, Saturation As Single, Luminance As Single
Static intNODE As Integer, YCtr As Integer, XCtr As Integer, intRainbowAngle As Integer
'There is no risk for getting dull shades since I use the native principal by adding/subtracting values against at constant FF-component.
'The algoritm gives med decimal values which increases the importance for mathematical models for choosing color, not pic.point.
'picThinBox.ScaleMode = vbPixels
intRainbowAngle = 0 'Protecting systemcolorangel
Saturation = 255: Luminance = 255 'Fully shining colors.
'If blnFadeToGrey = True And blnFadeToBlack = False Then Saturation = 255: Luminance = bteBrightnessMax255 'Starting value is full saturation. Brightness is to be the same for the whole bigbox.
'Horizontal or vertical kan be chosen by intKoordSuperior/intKoordInferior.
'YCtr = 255: If XCtr = YCtr Then blnHorizontal = True

bteK4243 = 42 'Has to alternate between 42 and 43 pixels per colorfield to make even at 256 pixels.

'Vertical
For XCtr = 0 To 19
    Do 'Y loopen 255 To 0.
    '1 Red in in direction towards yellow. Green is counting up.
    For Ctr = 1 To bteK4243  'Has to alternate between 42 and 43 pixels per colorfield to make even at 256 pixels.
        intRainbowAngle = intNODE + ((254 * (Ctr - 1)) / (bteK4243 - 1)) 'Wonderful solution: this logic about going from zero to the full value (here 254) I have been seeking for a long time.
        'objAnyPictureBox.PSet (XCtr, YCtr), HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, blnUpdatetextBoxes)
        SetPixelV picThinBox.hDC, XCtr, YCtr, HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, blnUpdateTextBoxes)
        YCtr = YCtr - 1
    Next Ctr '
    If bteK4243 = 43 Then bteK4243 = 42 Else bteK4243 = 43
    intNODE = intNODE + 255 'Bistabile switch.
    Loop While YCtr > 0
    intRainbowAngle = 0 'Painting the last fully red which is outside the logic of the routine.
    SetPixelV picThinBox.hDC, XCtr, YCtr, HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, blnUpdateTextBoxes)
    intNODE = 0
    YCtr = 255
    Next XCtr

End Sub

Private Sub picThinBox_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Y As Integer, intDirektion As Integer

If objOption(0) Then
    If KeyCode = vbKeyUp Then
        intDirektion = 1
        Call NudgeHueValue(ByVal intDirektion)
        Call picBigBox_Colorize
    End If
    If KeyCode = vbKeyDown Then
        intDirektion = -1
        Call NudgeHueValue(ByVal intDirektion)
        Call picBigBox_Colorize
    End If
End If
If objOption(1) Then '******
    MsgBox "Add code for radio1! Probably just writing in textbox Saturation!"
End If
If objOption(2) Then '*****
    MsgBox "Add code for radio2!"
End If

End Sub
Public Sub NudgeHueValue(ByVal intDirektion)
'1530 levels. The triangels are moving every sixth step and are lying on the byte level of 1530/6.
'RGBtxtboxes tells the nudge level:
Dim lngColor  As Long
'NudgeValue goes from ZERO to 1536.
intSystemColorAngleMax1530 = intSystemColorAngleMax1530 + intDirektion 'Calculating the new value of intSystemColorAngleMax1530, thus +1 or -1.

If intSystemColorAngleMax1530 > 1530 Then intSystemColorAngleMax1530 = 1530 'Limiter.
If intSystemColorAngleMax1530 < 0 Then intSystemColorAngleMax1530 = 0

lngColor = HSLToRGB(ByVal intSystemColorAngleMax1530, bteSaturationMax255, bteBrightnessMax255, True) 'lngColor as a function of HSLToRGB. System constants are being updated at the same time.
Call TriangelMove(255 - (intSystemColorAngleMax1530 / 1530 * 255))  'Moving the triangel.

End Sub

Public Sub SampleMarkerBackground()
Dim CtrX As Byte, CtrY As Byte
'Saving the background behind Marker.
If blnNotFirstTimeMarker = True Then
    For CtrY = 0 To 10
    For CtrX = 0 To 10
    arLongMarkerColorStore(CtrX, CtrY) = picBigBox.POINT(mBteMarkerOldX - 5 + CtrX, mBteMarkerOldY - 5 + CtrY)
    Next CtrX
    Next CtrY
End If
End Sub

Public Sub PaintMarker(X, Y)

    If bteBrightnessMax255 < 200 Then 'White marker if the surroundings are grey.
        picBigBox.Circle (X, Y), 5, vbWhite
        Exit Sub
    End If

    If Text1(0) < 26 Or Text1(0) > 200 Then 'Shades of blue.
        If bteSaturationMax255 > 70 Then ' And bteSaturationMax255 < 150 Then 'White marker if the surroundings are grey..
            picBigBox.Circle (X, Y), 5, vbWhite
            Exit Sub
        End If
    End If
    picBigBox.PaintPicture imgMarker, X - 5, Y - 5, 11, 11, 0, 0, 11, 11, vbSrcInvert 'Complementary colors
    picBigBox.PaintPicture imgMarker, X - 5, Y - 5, 11, 11, 0, 0, 11, 11, vbDstInvert

End Sub

Public Sub SplitlblNewColorToRGBboxes() 'Updating the system constants and textboxes regarding to RGB.
mSngRValue = lblNewColor.BackColor And &HFF: Text1(3) = mSngRValue
mSngGValue = (lblNewColor.BackColor And &HFF00&) \ &H100&: Text1(4) = mSngGValue
mSngBValue = (lblNewColor.BackColor And &HFF0000) \ &H10000: Text1(5) = mSngBValue
End Sub

Private Function RGBToHSL201(ByVal RGBValue As Long, ByVal blnUpdateTextBoxes As Boolean) As HSL
Dim R As Long, G As Long, B As Long
Dim lMax As Long, lMin As Long, lDiff As Long, lSum As Long

R = RGBValue And &HFF&
G = (RGBValue And &HFF00&) \ &H100&
B = (RGBValue And &HFF0000) \ &H10000

If R > G Then lMax = R: lMin = G Else lMax = G: lMin = R 'Finds the Superior and inferior components.
If B > lMax Then lMax = B Else If B < lMin Then lMin = B

lDiff = lMax - lMin
lSum = lMax + lMin
'Luminance, thus brightness' Adobe photoshop uses the logic that the site VBspeed regards (regarded) as too primitive = superior decides the level of brightness.
RGBToHSL201.Luminance = lMax / 255 * 100
'Saturation******
If lMax <> 0 Then 'Protecting from the impossible operation of division by zero.
    RGBToHSL201.Saturation = 100 * lDiff / lMax 'The logic of Adobe Photoshops is this simple.
Else
    RGBToHSL201.Saturation = 0
End If
'Hue ************** R is situated at the angel of 360 eller noll degrees; G vid 120 degrees; B vid 240 degrees. intSystemColorAngleMax1530
Dim q As Single
If lDiff = 0 Then q = 0 Else q = 60 / lDiff 'Protecting from the impossible operation of division by zero.
Select Case lMax
    Case R
        If G < B Then
            RGBToHSL201.Hue = 360& + q * (G - B)
        intSystemColorAngleMax1530 = (360& + q * (G - B)) * 4.25 'Converting from degrees to my resolution of detail.
        Else
            RGBToHSL201.Hue = q * (G - B)
        intSystemColorAngleMax1530 = (q * (G - B)) * 4.25
        End If
    Case G
        RGBToHSL201.Hue = 120& + q * (B - R) ' (R - G)
    intSystemColorAngleMax1530 = (120& + q * (B - R)) * 4.25
    Case B
        RGBToHSL201.Hue = 240& + q * (R - G)
    intSystemColorAngleMax1530 = (240& + q * (R - G)) * 4.25
End Select 'The case of B was missing.

If blnUpdateTextBoxes = True Then
    If R < &H10 Then
        txtHexColor = Right$("00000" & Hex$(R * 65536 + G * 256 + B), 6) 'Adds letters of zero to the left which is a necessary so called padding.
    Else
        txtHexColor = Hex$(R * 65536 + G * 256 + B)
    End If
    txtHexColor.Refresh 'End of hexabox routine.
    Text1(0) = Round(intSystemColorAngleMax1530 / 1530 * 360)
    If lMax = 0 Then
    bteSaturationMax255 = 0 'Protecting from the impossible operation of division by zero.
    Else
    bteSaturationMax255 = 255 * lDiff / lMax
    Text1(1) = RGBToHSL201.Saturation '= saturation both 0 To 255 and 0 To 100%.
    End If
    bteBrightnessMax255 = lMax: Text1(2) = RGBToHSL201.Luminance '=Brighness both 0 To 255 and 0 To 100%.
End If
End Function
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) Then 'Limiting the numerical textboxes (Text1[x]) to just register numerical enters.
        KeyAscii = 0
    End If
End Sub

 Sub Text1_LostFocus(Index As Integer)
Dim udtAngelSaturationBrightness As HSL, lngColor As Long 'Has to take care of intSystemColorAngleMax1530 0 To 1529.
    mBlnBigBoxReady = False 'Gives me fresh coordinates, but only in the RBG-model at this stage.
    blnNotFirstTimeMarker = False '-"-

'HAVE TO ADD THE FUNCTIONALITY: img.Pilars position is totally dependent of the actual mode.
If Index = 0 Then 'The user adjusted Hue so RGB will be aproximately calculated.
    If Text1(0) > 360 Then MsgBox "An integer between 0 and 360 i required. Closest value inserted!", vbCritical, "Color Picker": Text1(0) = 360 'Checking both the precense of decimals and numbers greater than 360.
    If Text1(0) <> Round(Text1(0)) Then MsgBox "An integer between 0 and 360 i required. Closest value inserted!", vbCritical, "Color Picker": Text1(0) = Round(Text1(0))  'Checking both the precense of decimals and numbers greater than 360.
    
    lngColor = HSLToRGB(Text1(0) / 360 * 255 * 6, bteSaturationMax255, bteBrightnessMax255, True)
End If
If Index = 1 Then 'The user adjusted Saturation so RGB will be aproximately calculated.
    If Text1(1) > 100 Then MsgBox "An number between 0 and 100 i required. Closest value inserted!", vbCritical, "Color Picker": Text1(1) = 100 'Checking both the precense of decimals and numbers greater than 360.
    If Text1(1) < 0 Then MsgBox "An number between 0 and 100 i required. Closest value inserted!", vbCritical, "Color Picker": Text1(1) = 0 'Checking both the precense of decimals and numbers greater than 360.
    lngColor = HSLToRGB(intSystemColorAngleMax1530, Text1(1) / 100 * 255, bteBrightnessMax255, True)
End If
If Index = 2 Then 'The user adjusted Luminance so RGB will be aproximately calculated.
    If Text1(2) > 100 Then MsgBox "An number between 0 and 100 i required. Closest value inserted!", vbCritical, "Color Picker": Text1(2) = 100 'Checking both the precense of decimals and numbers greater than 360.
    If Text1(2) < 0 Then MsgBox "An number between 0 and 100 i required. Closest value inserted!", vbCritical, "Color Picker": Text1(2) = 0 'Checking both the precense of decimals and numbers greater than 360.
    lngColor = HSLToRGB(intSystemColorAngleMax1530, bteSaturationMax255, Text1(2) / 100 * 255, True)
End If


If Index > 2 Then 'The user adjusted RGB so HSL is to calculated aproximately.
    udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True)
End If

'Justera imgArrows beroende på modus.
Call imgArrowsModeDepending

Call picBigBox_Colorize

End Sub
Private Sub txtHexColor_KeyPress(KeyAscii As Integer) 'Limits the textbox to numerics and A-F and capitals and to six pieces of letters.
'The limitation to six letters probably has to be done in the vb editor poreperties window (see Greg Perry VB in 6 days).
If (KeyAscii > 64 And KeyAscii < 71) Then Exit Sub 'A-F are OK.
If (KeyAscii > 96 And KeyAscii < 103) Then KeyAscii = KeyAscii - 32: Exit Sub 'a-f becomes A-F. OK.
If (KeyAscii > 47 And KeyAscii < 58) Then Exit Sub 'Numerics are OK.

KeyAscii = 0 'All other letters are unwanted.
End Sub

Private Sub txtHexColor_LostFocus()
'Dim udtAngelSaturationBrightness As HSL, lngColor As Long 'Must take care of intSystemColorAngleMax1530 0 To 1529.
'On Error GoTo Bajs
Dim sShift As String 'OBS! Must shift RGB into BGR to fit vb-standard.
sShift = txtHexColor: sShift = Mid(sShift, 5) & Mid(sShift, 3, 2) & Mid(sShift, 1, 2) 'Shifting RGB to BGR.
lblNewColor.BackColor = ("&H" + sShift) 'OBS! Must shift RGB into BGR to fit vb-standard.
Call SplitlblNewColorToRGBboxes 'Automatic update of the RGB textboxes.
Call Text1_LostFocus(3) 'Simulating that the user adjusted the RGBtxtboxes->Total update. 3 means that the RedTextbox has been adjusted.
Exit Sub
Bajs:
MsgBox "An error occured while translating hexnumber to decimal number!"
End Sub
Public Function HSLToRGB(ByVal intLocalColorAngle As Integer, ByVal Saturation As Long, ByVal Luminance As Long, ByVal blnUpdateTextBoxes As Boolean) As Long
Dim R As Long, G As Long, B As Long, lMax As Byte, lMid As Byte, lMin As Long, q As Single
lMax = Luminance
lMin = (255 - Saturation) * lMax / 255 '255 - (Saturation * lMax / 255)
q = (lMax - lMin) / 255

Select Case intLocalColorAngle
    Case 0 To 255
        lMid = (intLocalColorAngle - 0) * q + lMin
        R = lMax: G = lMid: B = lMin
    Case 256 To 510 'This period surpasses the node border with one unit - over to gren color. CHECK by F8.
        lMid = -(intLocalColorAngle - 255) * q + lMax '-(intLocalColorAngle - 256) * q + lMin
        R = lMid: G = lMax: B = lMin
    Case 511 To 765
        lMid = (intLocalColorAngle - 510) * q + lMin
        R = lMin: G = lMax: B = lMid
    Case 766 To 1020
        lMid = -(intLocalColorAngle - 765) * q + lMax
        R = lMin: G = lMid: B = lMax
    Case 1021 To 1275
        lMid = (intLocalColorAngle - 1020) * q + lMin
        R = lMid: G = lMin: B = lMax
    Case 1276 To 1530
        lMid = -(intLocalColorAngle - 1275) * q + lMax
        R = lMax: G = lMin: B = lMid
    Case Else
        MsgBox "Error occured in HSLToRGB. intSystemColorAngleMax1530= " & Str(intLocalColorAngle)
End Select

mSngRValue = R: mSngGValue = G: mSngBValue = B 'Updating the sustem constants automatically. Perhaps must exclude this to give them protection.
HSLToRGB = B * &H10000 + G * &H100& + R 'Delivers lngColor in VB-format.

If blnUpdateTextBoxes = True Then 'Then the calling routine is not any of the complex automatic routines for fading etc.
'Since this is a single time called session I can safely update my system constants and convert my hifgh resolution system constants to textbox dito.
    Text1(0) = Round(intLocalColorAngle / 255 / 6 * 360)
    Text1(1) = Round(Saturation / 255 * 100)
    Text1(2) = Round(Luminance / 255 * 100)
    Text1(3) = mSngRValue
    Text1(4) = mSngGValue
    Text1(5) = mSngBValue
    Text1(0).Refresh
    Text1(1).Refresh
    Text1(2).Refresh
    Text1(3).Refresh
    Text1(4).Refresh
    Text1(5).Refresh
    If mSngRValue < &H10 Then
        txtHexColor = Right$("00000" & Hex$(mSngRValue * 65536 + mSngGValue * 256 + mSngBValue), 6) 'Padding with zeroletters to the left.
    Else
        txtHexColor = Hex$(mSngRValue * 65536 + mSngGValue * 256 + mSngBValue)
    End If
    txtHexColor.Refresh 'End of the Hexabox routine.
    lblNewColor.BackColor = HSLToRGB
    lblNewColor.Refresh
    intSystemColorAngleMax1530 = intLocalColorAngle 'Sometims there is only a mouse Y coordinate tha is delivered from the calling routinen.
    bteSaturationMax255 = Saturation
    bteBrightnessMax255 = Luminance
End If
End Function
Private Sub lblNewColor_Click()
Dim udtAngelSaturationBrightness As HSL ', lngColor As Long
    lblOldColor.BackColor = lblNewColor.BackColor
    mBlnBigBoxReady = False 'Delivers fresh coordinates, but only in the HSL-model at this stage.
    blnNotFirstTimeMarker = False '-"-
    udtAngelSaturationBrightness = RGBToHSL201(lblNewColor.BackColor, True) 'True means that HSL is updating both the textboxes and the system constants.
    
    If objOption(9) Then Exit Sub 'Bail if postcard view.
    
    Call TriangelMove(255 - (intSystemColorAngleMax1530 / 1530 * 255) + 28) 'Animates the triangeln.
    Call picBigBox_Colorize   'Redraw BigBox

End Sub
Private Sub lblOldColor_Click()
Dim udtAngelSaturationBrightness As HSL ', lngColor As Long
    lblNewColor.BackColor = lblOldColor.BackColor
    mBlnBigBoxReady = False 'Delivers fresh coordinates, but only in the HSL-model at this stage.
    blnNotFirstTimeMarker = False '-"-
    udtAngelSaturationBrightness = RGBToHSL201(lblNewColor.BackColor, True) 'True means that HSL is updating both the textboxes and the system constants.
    Call TriangelMove(255 - (intSystemColorAngleMax1530 / 1530 * 255) + 28) 'Animates the triangel.
    Call picBigBox_Colorize   'Rita om BigBox

End Sub
Private Sub imgArrowsModeDepending()
'AdjustingJusterar imgArrows depending on current mode.
If objOption(0) Then Call TriangelMove(255 - (intSystemColorAngleMax1530 / 1530 * 255)) 'Animating the triangel.
If objOption(1) Then Call TriangelMove(255 - (Text1(1) * 2.55))  'Animating the triangel.
If objOption(2) Then Call TriangelMove(255 - (Text1(2) * 2.55))  'Animating the triangel.
If objOption(3) Then Call TriangelMove(255 - Text1(3))   'Animating the triangel.

If objOption(9) Then linTriang1Vert.Visible = False: linTriang1Rising.Visible = False: linTriang1Falling.Visible = False 'Top = 255 - (Text1(2) * 2.55) + 28  'Animating the triangel.

End Sub

Private Sub lblComplementaryColor_Click(Index As Integer)
If Text1(0) < 180 Then Text1(0) = Text1(0) + 180 Else Text1(0) = Text1(0) - 180
Call Text1_LostFocus(0) 'Noll stands for Hue.
End Sub
Private Sub PaintThinBox(Index As Integer)
Dim blnFadeToGrey As Boolean, blnFadeToBlack As Boolean

If Index = 0 Then
    Call RainBowThinBox
End If

If Index = 1 Then
    Call FadeThinBoxToGrey
    picThinBox.Refresh
    End If

If Index = 2 Then ' "Brightness"
    picThinBox.BackColor = HSLToRGB(ByVal intSystemColorAngleMax1530, ByVal bteSaturationMax255, ByVal 255, False) 'Delivers a lighter shade of the active color. 'Setting the whole square for easy fading.
    Call FadeThinBoxToBlack
    picThinBox.Refresh
End If
picThinBox.Visible = True

End Sub
Private Sub TriangelMove(Y)
linTriang1Vert.Y1 = Y + 28: linTriang1Vert.Y2 = Y + 28 + 10
linTriang1Rising.Y1 = Y + 28 + 10: linTriang1Rising.Y2 = Y + 28 + 4
linTriang1Falling.Y1 = Y + 28: linTriang1Falling.Y2 = Y + 28 + 6

linTriang2Vert.Y1 = Y + 28: linTriang2Vert.Y2 = Y + 28 + 10
linTriang2Rising.Y2 = Y + 28 + 10: linTriang2Rising.Y1 = Y + 28 + 5
linTriang2Falling.Y2 = Y + 28: linTriang2Falling.Y1 = Y + 28 + 5

End Sub
Public Sub opt3RedPaintPicThinBox(ByVal G, B)
Dim bteX As Byte, intCtr As Integer
For bteX = 0 To 19
For intCtr = 0 To 255
    SetPixelV picThinBox.hDC, bteX, intCtr, RGB(255 - intCtr, G, B) 'Painting with API.
Next intCtr
Next bteX

End Sub
Public Sub opt4GreenPaintPicThinBox(ByVal R, B)
Dim bteX As Byte, intCtr As Integer
For bteX = 0 To 19
For intCtr = 0 To 255
    SetPixelV picThinBox.hDC, bteX, intCtr, RGB(R, 255 - intCtr, B) 'Painting by API.
Next intCtr
Next bteX

End Sub
Public Sub opt5BluePaintPicThinBox(ByVal R, G)
Dim bteX As Byte, intCtr As Integer

For bteX = 0 To 19
For intCtr = 0 To 255
    SetPixelV picThinBox.hDC, bteX, intCtr, RGB(R, G, 255 - intCtr) 'Painting by API.
Next intCtr
Next bteX

End Sub

Public Sub BigBoxOpt3Reaction(ByVal X, Y)
Dim udtAngelSaturationBrightness As HSL

lblNewColor.BackColor = picBigBox.POINT(X, Y): lblNewColor.Refresh
    Call SplitlblNewColorToRGBboxes 'Updating the module global mSngRValue etc.
    Call opt3RedPaintPicThinBox(ByVal mSngGValue, mSngBValue)
    picThinBox.Refresh
    udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True) 'Causes an update of all the constants.
End Sub
Public Sub BigBoxOpt4Reaction(ByVal X, Y)
Dim udtAngelSaturationBrightness As HSL

lblNewColor.BackColor = picBigBox.POINT(X, Y): lblNewColor.Refresh
    Call SplitlblNewColorToRGBboxes 'Updating the module global mSngRValue etc.
    Call opt4GreenPaintPicThinBox(ByVal mSngRValue, mSngBValue)
    picThinBox.Refresh
    udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True) 'Causes an update of all the constants. The letter of three stands for the RED-box.
End Sub
Public Sub BigBoxOpt5Reaction(ByVal X, Y)
Dim udtAngelSaturationBrightness As HSL

lblNewColor.BackColor = picBigBox.POINT(X, Y): lblNewColor.Refresh
    Call SplitlblNewColorToRGBboxes 'Updating the module global mSngRValue etc.
    Call opt5BluePaintPicThinBox(ByVal mSngRValue, mSngGValue)
    picThinBox.Refresh
    udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True) 'Causes an update of all the constants. The letter of three stands for the RED-box..
End Sub
Private Sub opt3RedPaintPicBigBox()
Dim R As Single, G As Single, B As Single
R = Text1(3) 'Red
For B = 255 To 0 Step -1
For G = 255 To 0 Step -1 'Interesting if there is an error, thus a jump directly to EndSub.
    SetPixelV picBigBox.hDC, B, 255 - G, RGB(R, G, B) 'Painting by API.
Next G
Next B

End Sub
Private Sub opt4GreenPaintPicBigBox()
Dim R As Single, G As Single, B As Single
G = Text1(4) 'Green
For B = 255 To 0 Step -1
For R = 255 To 0 Step -1 'Interesting if there is an error, thus a jump directly to EndSub.
    SetPixelV picBigBox.hDC, B, 255 - R, RGB(R, G, B) 'Painting by API.
Next R
R = R - 1 'Because that R becomes too big when the loop has finishes.
Next B

End Sub
Private Sub opt5BluePaintPicBigBox()
Dim R As Single, G As Single, B As Single
B = Text1(5) 'Blue
For R = 255 To 0 Step -1
For G = 255 To 0 Step -1 'Interesting if there is an error, thus a jump directly to EndSub.
    SetPixelV picBigBox.hDC, R, 255 - G, RGB(R, G, B) 'Ritar medelst API.
Next G
G = G - 1 'Because that G becomes too big when the loop has finishes..
Next R

End Sub
Public Sub ExecuteIniFile(bteValdRadioKnapp)
Dim bteFileHandle As Byte, Ctr As Byte, strTillfRadioKnapp As String, strTillfNumberOfPaths As String, bteNumberOfPaths As Byte
Dim strColor As String, strFilename As String, Answer As Integer, Result As Integer, sFile As String, strAppPath As String
'Interna bilder to add to combo box:
Combo1.AddItem "Winterlake.jpg"
Combo1.AddItem "50's Colormap.jpg"
'To remove from a combo box:

'Opens too read from the ini-file of the form.
bteFileHandle = FreeFile
On Error GoTo ErrorHandler 'Error 53 means that the SPECIAL file handling routines cant find the file.
strAppPath = App.Path
If InStrRev(strAppPath, "\") <> Len(strAppPath) Then strAppPath = strAppPath & "\"  'Good windows-programming-manners.
Open strAppPath & "Colorpicker.ini" For Input As bteFileHandle
Line Input #bteFileHandle, strColor
If IsNumeric(Trim(strColor)) Then lblOldColor.BackColor = strColor 'Trim removes space och citation marks on either side.
Line Input #bteFileHandle, strTillfRadioKnapp
Line Input #bteFileHandle, strTillfNumberOfPaths
If IsNumeric(Mid(strTillfNumberOfPaths, Len(strTillfNumberOfPaths) - 1, 1)) Then 'Safety precaution.
    bteNumberOfPaths = Mid(strTillfNumberOfPaths, Len(strTillfNumberOfPaths) - 1, 1)
End If
ReDim arsPicPath(bteNumberOfPaths)

If bteNumberOfPaths = 0 Then
    ReDim arsPicPath(1): arsPicPath(1) = "" 'A flag of the save routine.
Else
    For Ctr = 1 To bteNumberOfPaths
        Line Input #bteFileHandle, arsPicPath(Ctr): arsPicPath(Ctr) = Mid(arsPicPath(Ctr), 2, Len(arsPicPath(Ctr)) - 2) 'Removing the citation marks - the Trim command wasn't sufficient.
        Combo1.AddItem Mid(arsPicPath(Ctr), InStrRev(arsPicPath(Ctr), "\") + 1)  'Extracting the file name from the pathen, but with extension ".wav".
    Next Ctr
End If

Close #bteFileHandle 'Protecting the file.
If IsNumeric(Mid(strTillfRadioKnapp, Len(strTillfRadioKnapp) - 1, 1)) Then 'Precaution.
    bteValdRadioKnapp = Mid(strTillfRadioKnapp, Len(strTillfRadioKnapp) - 1, 1) 'There is a number between 0-9.
End If
Exit Sub

ErrorHandler: 'In case file is missing
If Err.Number <> 53 And Err.Number <> 76 Then '53 means FileNotFound, 76 means "path not found".
    MsgBox "Unknown error nr " & Err.Number & Err.Description: Exit Sub
Else
'53 and 76 are error-numbers for missing file by the special file handling routines.
Answer = MsgBox("Click YES to create your backup file! In case this is not your first launch of ColorPicker the file Colorpicker.ini which contains your last settings is lost! Then you have the option to click NO to manually search for the file on your drive?", vbQuestion + vbYesNo, "Welcome first time user - there is no log-file yet!")
    If Answer = vbNo Then
        For Ctr = 1 To 9
            If Ctr > 8 Then MsgBox "You have tried 8 times and could probably be stuck in some disfunctionality. I will bail you out. Please start the programme again!": End
            Call OpenDialog(sFile)
            'Extracting the picname + extension from the path.
            strFilename = Mid(sFile, InStrRev(sFile, "\") + 1) 'Extracting the file name from the pathen, but with extension ".wav"..
            If strFilename = "Colorpicker.ini" Then Exit For
            If strFilename <> "Colorpicker.ini" Then
                Result = MsgBox("The file name must be Colorpicker.ini! Do you want to retry? The answer No will create a new blank ini-file when leaving the application!", vbQuestion + vbYesNo, "ini-file missing!")
                If Answer = vbYes Then ReDim arsPicPath(1): arsPicPath(1) = "": Exit Sub 'Flag to the save-routine.
            End If
        Next Ctr
        
        MsgBox "I will now move the ini file to it's correct location in the directory colorpicker" 'Putting the correct path in the matrix.
        Name sFile As strAppPath & "Colorpicker.ini": Exit Sub 'Moving the imagefile to its correct location.
    End If
End If
ReDim arsPicPath(1): arsPicPath(1) = "" 'Is executed if the user refused to manually search for the ini-file.
End Sub
Private Sub MoveHexBox()
Dim Ctr As Integer
For Ctr = 336 To 286 Step -1
txtHexColor.Move Ctr, 281, 56, 20
Combo1.Move Ctr + 70, 281, 70 + 336 - Ctr 'Height-property in ComboBoxes is readonly.
Next Ctr

End Sub
Public Sub RepairLink(sFile)
Dim Ctr As Byte
Call OpenDialog(sFile)
If sFile = "" Then sFile = "Cancel": Exit Sub 'User pressed cancel. Using the sFile as a messenger.
arsPicPath(Combo1.ListIndex - 1) = sFile 'Putting the correct path in the matrix.
lblPicPath = sFile

'At this stage the whole Combobox must be emptied and rebuilt from the matrix.
Combo1.Clear
Combo1.AddItem "Winterlake.jpg"
Combo1.AddItem "50's Colormap.jpg"

For Ctr = 1 To UBound(arsPicPath)
Combo1.AddItem Mid(arsPicPath(Ctr), InStrRev(arsPicPath(Ctr), "\") + 1)   'Loading with the filename only.
Next Ctr
End Sub
Public Sub RemoveBrokenLinks()
Dim Ctr As Byte, RetVal As String, bteListIndex As Byte
Ctr = 1
Do
RetVal = Dir(arsPicPath(Ctr)) 'The function of Dir returns the path if the file exists, otherwise an empty string will be returned.
If RetVal <> Mid(arsPicPath(Ctr), InStrRev(arsPicPath(Ctr), "\") + 1) Then 'Extracting the filename from the path.
    bteListIndex = Ctr + 1 'The pointer of the Combobox (which by the way has the base of zero) and two native images.
    Call ContractArrays(bteListIndex) 'A broken link to remove! Reusing some code.
    Ctr = Ctr - 1 'The list has been decremented by one step.
End If
Ctr = Ctr + 1
Loop While Ctr < UBound(arsPicPath) + 1

End Sub
Public Sub ContractArrays(bteListIndex) 'Contracting by one cell unit.
Dim Ctr As Byte
For Ctr = bteListIndex - 1 To UBound(arsPicPath) - 1 'Counting in the matrix (which has the base of one) from the highlighted post to the second last post.
    arsPicPath(Ctr) = arsPicPath(Ctr + 1) 'Contracting the information.
Next Ctr
Combo1.RemoveItem (bteListIndex): bteListIndex = bteListIndex - 1 'The base of the list is noll.
ReDim Preserve arsPicPath(UBound(arsPicPath) - 1) 'Contracting the matrix by one cell.
End Sub

Private Sub OpenDialog(sFile)
With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.Filename) = 0 Then
            Exit Sub
        End If
        sFile = .Filename
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim bteFileHandle As Byte, Ctr As Byte, strAppPath As String
    
    For Ctr = 0 To 9 'Determing which optRadioknapp that was chosen the last session.
        If objOption(Ctr) Then Exit For
    Next Ctr

    'Opening to write to the ini-file of the form.
    bteFileHandle = FreeFile

Exit Sub

ErrorHandler:
    MsgBox "Unknown error" & Err.Description: End 'Bail
End Sub

