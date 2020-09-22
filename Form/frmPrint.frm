VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Preview"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9165
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmOrientation 
      Caption         =   "Orientation"
      Height          =   2115
      Left            =   6345
      TabIndex        =   12
      Top             =   1560
      Width           =   2640
      Begin VB.PictureBox picOrientation 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   150
         ScaleHeight     =   645
         ScaleWidth      =   2475
         TabIndex        =   15
         Top             =   300
         Width           =   2475
         Begin VB.OptionButton optOrien 
            Caption         =   "Landscape"
            Height          =   255
            Index           =   1
            Left            =   765
            TabIndex        =   17
            Top             =   345
            Width           =   1590
         End
         Begin VB.OptionButton optOrien 
            Caption         =   "Portrait"
            Height          =   255
            Index           =   0
            Left            =   795
            TabIndex        =   16
            Top             =   0
            Width           =   1590
         End
         Begin VB.Image imgPage 
            Height          =   345
            Index           =   1
            Left            =   435
            Picture         =   "frmPrint.frx":0000
            Top             =   135
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Image imgPage 
            Height          =   465
            Index           =   0
            Left            =   30
            Picture         =   "frmPrint.frx":0585
            Top             =   75
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Image imgPrinterOrien 
            Height          =   465
            Left            =   195
            Top             =   105
            Width           =   345
         End
      End
      Begin VB.OptionButton optPaperSize 
         Caption         =   "A4"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   14
         Top             =   1560
         Width           =   540
      End
      Begin VB.OptionButton optPaperSize 
         Caption         =   "8.5 x 11"""
         Height          =   195
         Index           =   1
         Left            =   1245
         TabIndex        =   13
         Top             =   1575
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Paper size"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   1260
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   420
      Left            =   6975
      TabIndex        =   8
      Top             =   4470
      Width           =   1305
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   420
      Left            =   6975
      TabIndex        =   7
      Top             =   3975
      Width           =   1305
   End
   Begin VB.CheckBox chkVel 
      Appearance      =   0  'Flat
      Caption         =   "Custom Size"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6495
      TabIndex        =   6
      Top             =   300
      Width           =   1485
   End
   Begin VB.TextBox txtWidth 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   7245
      TabIndex        =   3
      Top             =   660
      Width           =   705
   End
   Begin VB.TextBox txtHeight 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   7245
      TabIndex        =   2
      Top             =   960
      Width           =   705
   End
   Begin VB.PictureBox picPaper 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   405
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   0
      Top             =   690
      Width           =   3150
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   1950
         Left            =   405
         MousePointer    =   15  'Size All
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   144
         TabIndex        =   11
         Top             =   615
         Width           =   2190
      End
      Begin VB.Image imgPreview 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1110
         Left            =   1605
         MousePointer    =   15  'Size All
         Stretch         =   -1  'True
         Top             =   2295
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Shape shpMargin 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Height          =   4230
         Left            =   60
         Top             =   60
         Width           =   3000
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Picture"
      Height          =   420
      Left            =   6975
      TabIndex        =   9
      Top             =   3975
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   11040
      Left            =   8385
      ScaleHeight     =   732
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   10
      Top             =   4890
      Visible         =   0   'False
      Width           =   15420
   End
   Begin VB.Image imgBuffer 
      Height          =   945
      Left            =   1365
      Top             =   4440
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Height"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   1
      Left            =   6525
      TabIndex        =   5
      Top             =   960
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Width"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   6525
      TabIndex        =   4
      Top             =   660
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "PRINT PREVIEW - A4 PAPER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   405
      TabIndex        =   1
      Top             =   315
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   165
      Top             =   180
      Width           =   6045
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(c) 2007 Diomidisk

Option Explicit
Private aMouseDown As Boolean
Private xp1 As Long, yp1 As Long
Private ratio As Single
Private Canceled As Boolean
Private W As Single
Private H As Single
Private LeftM As Single
Private TopM As Single
Private RightM As Single
Private BottomM As Single

Private Sub chkVel_Click()
      If chkVel.Value = 1 Then
         TxtWidth.Enabled = True
         TxtHeight.Enabled = True
      Else
        TxtWidth.Enabled = False
        TxtHeight.Enabled = False
      End If
End Sub

Private Sub cmdExit_Click()
 
   Canceled = False
    Hide
End Sub

Private Sub cmdOpen_Click()
   
        Dim W As Long, H As Long
        
        On Error Resume Next
        'Set imgBuffer.Picture to the picture from the file
        imgBuffer.Picture = Picture1.Picture 'LoadPicture(FileName)
        ratio = imgBuffer.Width / imgBuffer.Height
        
        'Put the image to scale according to paper size
        pic.Width = imgBuffer.Width / 2.8
        pic.Height = imgBuffer.Height / 2.8
        pic.Picture = imgBuffer.Picture
        
       
        'If the image is too wide resize it but constrain proportions
        'You should add similar code for height
        If pic.Left + pic.Width > shpMargin.Left + shpMargin.Width Then
            If pic.Width > 560 / 2.8 Then
                pic.Width = 560 / 2.8
                pic.Height = pic.Width / ratio
            End If
            pic.Move shpMargin.Left
        End If
        'Set resize labels
        TxtHeight.Text = Int(pic.Height * 2.8)
        TxtWidth.Text = Int(pic.Width * 2.8)
         W = pic.Width
         H = pic.Height
         pic = ResizeGDIPlus(Picture1.Picture.Handle, Picture1.Picture.Type, W, H, , True)
        cmdPrint.Enabled = True
        
        If Err Then
            MsgBox Err.Description, vbInformation, App.Title
        End If
    
End Sub

Private Sub cmdPrint_Click()
    
    
    If TxtWidth.Text > 0 Then
       
       pic.Width = ((pic.Width * 2.8) / 28) * 546.44
       pic.Height = ((pic.Height * 2.8) / 28) * 546.44
                            
'        Printer.PaintPicture pic.Picture, ((pic.Left * 2.8) / 28) * 546.44, ((pic.Top * 2.8) / 28) * 546.44, ((pic.Width * 2.8) / 28) * 546.44, ((pic.Height * 2.8) / 28) * 546.44
        Printer.PaintPicture Picture1.Picture, ((pic.Left * 2.8) / 28) * 546.44, ((pic.Top * 2.8) / 28) * 546.44, pic.Width, pic.Height

        Printer.EndDoc
    End If
    cmdExit_Click
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Unload Me
    
End Sub


Private Sub optOrien_Click(Index As Integer)
  Select Case Index
    Case 0
          imgPrinterOrien.Picture = imgPage(Index).Picture
'           myW = tmpW
'           myH = tmpH
           gPrintetOrientation = vbPRORPortrait
    Case 1
          imgPrinterOrien.Picture = imgPage(Index).Picture
'           myW = tmpH
'           myH = tmpW
           gPrintetOrientation = vbPRORLandscape
    End Select
    Printer.Orientation = gPrintetOrientation
    SetPage
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
aMouseDown = True
   xp1 = x
   yp1 = y
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim NewX As Long, NewY As Long

   If aMouseDown Then
      NewX = pic.Left + (x - xp1)
      NewY = pic.Top + (y - yp1)
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

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   aMouseDown = False

End Sub

Private Sub txtHeight_Change()
    
    On Error Resume Next
    
    If TxtHeight.Text = "" Then Exit Sub
    If Int(TxtHeight.Text) <= 792 Then
        pic.Height = Int(TxtHeight.Text) / 2.8
    Else
        TxtHeight.Text = "792"
    End If
    
    pic.Width = pic.Height * ratio
   '
    pic = ResizeGDIPlus(Picture1.Picture.Handle, _
                        Picture1.Picture.Type, _
                        pic.Width, _
                        pic.Height)
    imgBuffer.Picture = pic.Picture
    
    pic.PaintPicture imgBuffer.Picture, 0, 0, pic.Width, pic.Height
    
End Sub

Private Sub txtWidth_Change()

    On Error Resume Next
   
    If TxtWidth.Text = "" Then Exit Sub
    If Int(TxtWidth.Text) <= 560 Then
        If Int(TxtWidth.Text) > 0 Then
            pic.Width = Int(TxtWidth.Text) / 2.8
        End If
    Else
        TxtWidth.Text = "560"
    End If
    pic.Height = pic.Width / ratio
    pic = ResizeGDIPlus(Picture1.Picture.Handle, _
                        Picture1.Picture.Type, _
                        pic.Width, _
                        pic.Height)
    imgBuffer.Picture = pic.Picture
    pic.PaintPicture imgBuffer.Picture, 0, 0, pic.Width, pic.Height
    
End Sub

Sub ChangeSize()

        ratio = imgBuffer.Width / imgBuffer.Height
        pic.Height = pic.Width / ratio
        
        pic.PaintPicture imgBuffer.Picture, 0, 0, pic.Width, pic.Height
        TxtHeight.Text = Int(pic.Height * 2.8)
        TxtWidth.Text = Int(pic.Width * 2.8)
        
End Sub

' Display the form. Return True if the user cancels.
Public Function ShowForm(cPic As Picture)
       Set Picture1.Picture = cPic
       cmdOpen_Click
       SetPage
       Show vbModal
       Unload Me
End Function

Sub SetPage()
    
    W = Printer.ScaleX(Printer.Width, Printer.ScaleMode, vbMillimeters)
    H = Printer.ScaleY(Printer.Height, Printer.ScaleMode, vbMillimeters)
    picPaper.Width = W
    picPaper.Height = H
    
    picPaper.Move (Shape1.Width - picPaper.Width) / 2, (Shape1.Height - picPaper.Height) / 2 + Label1.Height
    
    If LeftM = 0 And TopM = 0 And RightM = 0 And BottomM = 0 Then
       LeftM = 5: TopM = 5: RightM = 5: BottomM = 5
    End If
    
    If gPrintetOrientation = 2 Then
       optOrien(1).Value = True
    Else
       optOrien(0).Value = True
    End If
    
    shpMargin.Move LeftM, TopM, W - RightM - LeftM, H - BottomM - TopM
    pic.Move (shpMargin.Width - pic.Width) / 2, (shpMargin.Height - pic.Height) / 2
    cmdOpen_Click
    'If pic.Height > shpMargin.Height Then
    '   ratio = imgBuffer.Width / imgBuffer.Height
    '   pic.Width = shpMargin.Width * ratio
    '   pic.Height = shpMargin.Height * ratio
    'End If
End Sub
