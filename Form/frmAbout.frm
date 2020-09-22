VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About MyApp"
   ClientHeight    =   4590
   ClientLeft      =   2340
   ClientTop       =   1815
   ClientWidth     =   6000
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3168.1
   ScaleMode       =   0  'User
   ScaleWidth      =   5634.31
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   225
      Top             =   1305
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2670
      Left            =   105
      ScaleHeight     =   174
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   382
      TabIndex        =   5
      Top             =   1785
      Width           =   5790
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1125
      Left            =   60
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   790.125
      ScaleMode       =   0  'User
      ScaleWidth      =   1053.5
      TabIndex        =   1
      Top             =   165
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4530
      TabIndex        =   0
      Top             =   1305
      Width           =   1260
   End
   Begin VB.Label lblDescription 
      Caption         =   "Draw application"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1620
      TabIndex        =   2
      Top             =   1065
      Width           =   2850
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   1650
      TabIndex        =   3
      Top             =   165
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   1635
      TabIndex        =   4
      Top             =   720
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(c) 2007-2010 - Dk
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Canceled As Boolean
Dim Scrolling As Boolean    'Scroll flag
Dim m_View As Boolean

Private Sub cmdOK_Click()
     Timer1.Enabled = False
     Canceled = False
     Scrolling = False
     Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    FormOnTop Me, True
 
End Sub

Private Sub Scroll(pic As PictureBox, TxtScroll As String, Optional Alignment As Long = &H1)
        
        Dim TextLine() As String    'Text lines array
        
       ' Dim Alignment As Long       'Text alignment
        Dim t As Long               'Timer counter (frame delay)
        Dim Index As Long           'Actual line index
        Dim RText As RECT           'Rectangle into each new text line will be drawed
        Dim RClip As RECT           'Rectangle to scroll up
        Dim RUpdate As RECT         'Rectangle to update (not used)
        
        If TxtScroll = "" Then Exit Sub
        TextLine() = Split(TxtScroll, vbCr)
        
        With pic
             .ScaleMode = vbPixels
             .AutoRedraw = True
            'Set rectangles
             SetRect RClip, 0, 1, .ScaleWidth, .ScaleHeight
             SetRect RText, 0, .ScaleHeight, .ScaleWidth, .ScaleHeight + .TextHeight("")
             
        End With
        
        Dim txt As String 'Text to be drawed
        With pic
            Do
               'Periodic frames
                If GetTickCount - t > 25 Then 'Set your delay here [ms]
                  'Reset timer counter
                   t = GetTickCount
                  'Line ( + spacing ) totaly scrolled ?
                   If RText.Bottom < .ScaleHeight Then
                     'Move down Text area out scroll area...
                      OffsetRect RText, 0, .TextHeight("") ' + space between lines [Pixels]
                     'Get new line
                      If Alignment = &H1 Then
                        'If alignment = Center, remove spaces
                         txt = Trim(TextLine(Index))
                      Else
                        'Case else, preserve them
                         txt = TextLine(Index)
                      End If
                     'Source line counter...
                      Index = Index + 1
                   End If
                  'Draw text
                   DrawText .hDC, txt, Len(txt), RText, Alignment
                  'Move up one pixel Text area
                   OffsetRect RText, 0, -1
                
                  'Finaly, scroll up (1 pixel)...
                   ScrollDC .hDC, 0, -1, RClip, RClip, 0, RUpdate

                  '...and draw a bottom line to prevent... (well, don't draw it and see what happens)
                   pic.Line (0, .ScaleHeight - 1)-(.ScaleWidth, .ScaleHeight - 1), .BackColor
                  '(Refresh doesn't needed: any own PictureBox draw method calls Refresh method)
                 End If
                 DoEvents
                
            Loop Until Scrolling = False Or Index > UBound(TextLine)
        End With
       If m_View = False Then
          Unload Me
        Else
          If Scrolling Then Timer1_Timer
       End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
       'Cancel = -1
       'cmdOK_Click
       Canceled = False
     Scrolling = False
End Sub

Private Sub Timer1_Timer()
     Dim txt As String
     txt = "ArtDraw " + vbCr + vbCr + lblVersion.Caption + vbCr + vbCr + "(c) 2007-2010 Diomidis Kiriakopoulos" + vbCr + vbCr + " Special Thanks:" + vbCr + _
           "Rod Stephens (vb-helper.com)" + vbCr + _
           "www.planet-source-code.com" + vbCr + "www.vbaccelerator.com" + vbCr + "vb-helper.com" + vbCr + _
           "edais.mvps.org" + vbCr + "www.codeguru.com" + vbCr + "binaryworld.net" + vbCr + vbCr + vbCr + vbCr + vbCr + vbCr '+ _
           '"Ron van Tilburg - psc" + vbCr + _
           "Carles P.V. - psc" + vbCr + _
           "Anna Carin - psc" + vbCr + vbCr + _
           "And for graphic filter" + vbCr + "Manuel Augusto Nogueira" + vbCr + "" + vbCr + vbCr + vbCr
     Timer1.Enabled = False
     Scrolling = True
     Scroll Picture1, txt
End Sub


' Display the form. Return True if the user cancels.
Public Function ShowForm(Optional sView As Boolean = False) As Boolean
    m_View = sView
    cmdOK.Visible = sView
    Timer1.Enabled = True
    ' Display the form.
    Show vbModal
    ShowForm = Canceled
    Unload Me
End Function


