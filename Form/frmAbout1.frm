VERSION 5.00
Begin VB.Form frmAbout1 
   BorderStyle     =   0  'None
   ClientHeight    =   5445
   ClientLeft      =   2295
   ClientTop       =   1500
   ClientWidth     =   7260
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3758.236
   ScaleMode       =   0  'User
   ScaleWidth      =   6817.515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5490
      Left            =   15
      Picture         =   "frmAbout1.frx":0000
      ScaleHeight     =   5430
      ScaleWidth      =   7200
      TabIndex        =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   7260
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "App Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2745
      TabIndex        =   3
      Top             =   2340
      Width           =   1365
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2295
      TabIndex        =   2
      Top             =   1935
      Width           =   930
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Art Draw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Left            =   1320
      TabIndex        =   1
      Top             =   1425
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private m_cDibR As New cDIBSectionRegion

Private Const AviDownload As Long = 104

Private Sub Form_Load()
    OnForm Me, True
    'Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblTitle.Caption = App.Title
   lblDescription.Caption = App.LegalCopyright
         
   Dim cDib As New cDIBSection
   cDib.CreateFromPicture Picture1.Picture   'picBackground.Picture
   m_cDibR.Create cDib
   m_cDibR.Applied(Me.hWnd) = True
   Set Me.Picture = Picture1.Picture   'picBackground.Picture
   Me.Show
   
  ' lblDescription.Left = (Picture1.Width - lblDescription.Width) / 2
  ' lblVersion.Left = (Picture1.Width - lblVersion.Width) / 2
       
   Me.Refresh
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' A better (more flexible) way of doing this is to use
   ' the vbAccelerator WM_NCHITTEST interception library.
   ' but if you want minimal code, here is the quick way!
    If Button = vbLeftButton Then
        'Fake a mouse down on the titlebar so form can be moved...
        ReleaseCapture
        SendMessageLong Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Unload Me
End Sub


Private Sub LblLoad_Change()
Me.Refresh
End Sub


