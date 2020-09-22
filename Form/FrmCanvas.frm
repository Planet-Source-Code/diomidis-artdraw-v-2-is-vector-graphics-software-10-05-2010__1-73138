VERSION 5.00
Begin VB.Form FrmCanvas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Page Size"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7245
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   5265
      TabIndex        =   20
      Top             =   3420
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton OptionColor 
      Caption         =   "Color"
      Height          =   210
      Left            =   5250
      TabIndex        =   19
      Top             =   3165
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Frame frmOrientation 
      Caption         =   "Orientation"
      Height          =   1200
      Left            =   3675
      TabIndex        =   15
      Top             =   1650
      Width           =   3330
      Begin VB.PictureBox picOrientation 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   105
         ScaleHeight     =   645
         ScaleWidth      =   3015
         TabIndex        =   16
         Top             =   270
         Width           =   3015
         Begin VB.OptionButton optOrien 
            Caption         =   "Portrait"
            Height          =   255
            Index           =   0
            Left            =   930
            TabIndex        =   18
            Top             =   0
            Value           =   -1  'True
            Width           =   1590
         End
         Begin VB.OptionButton optOrien 
            Caption         =   "Landscape"
            Height          =   255
            Index           =   1
            Left            =   945
            TabIndex        =   17
            Top             =   345
            Width           =   1590
         End
         Begin VB.Image imgPage 
            Height          =   345
            Index           =   1
            Left            =   405
            Picture         =   "FrmCanvas.frx":0000
            Top             =   195
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Image imgPage 
            Height          =   465
            Index           =   0
            Left            =   0
            Picture         =   "FrmCanvas.frx":0585
            Top             =   135
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
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4020
      TabIndex        =   14
      Top             =   4020
      Width           =   1080
   End
   Begin VB.OptionButton OptionType 
      Caption         =   "Millimeters"
      Height          =   240
      Index           =   2
      Left            =   5550
      TabIndex        =   13
      Top             =   990
      Width           =   1125
   End
   Begin VB.OptionButton OptionType 
      Caption         =   "Inches"
      Height          =   240
      Index           =   1
      Left            =   5550
      TabIndex        =   12
      Top             =   720
      Width           =   1125
   End
   Begin VB.OptionButton OptionType 
      Caption         =   "Pixels"
      Height          =   240
      Index           =   0
      Left            =   5550
      TabIndex        =   11
      Top             =   450
      Width           =   1110
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   150
      TabIndex        =   6
      Top             =   390
      Width           =   3330
   End
   Begin VB.CommandButton ComOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2505
      TabIndex        =   0
      Top             =   4020
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "Custom"
      Height          =   1365
      Left            =   3675
      TabIndex        =   1
      Top             =   165
      Width           =   3330
      Begin VB.TextBox TxtHeight 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "480"
         Top             =   675
         Width           =   690
      End
      Begin VB.TextBox TxtWidth 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "640"
         Top             =   285
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Height:"
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Width:"
         Height          =   285
         Left            =   195
         TabIndex        =   4
         Top             =   330
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Backcolor Page"
      Height          =   960
      Left            =   3705
      TabIndex        =   8
      Top             =   2835
      Width           =   3315
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   165
         Picture         =   "FrmCanvas.frx":0B33
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Select page color"
         Top             =   270
         Width           =   495
      End
      Begin VB.PictureBox PicColor 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   690
         ScaleHeight     =   390
         ScaleWidth      =   705
         TabIndex        =   10
         Top             =   285
         Width           =   765
      End
   End
   Begin VB.Label LabelSize 
      Caption         =   "Size :"
      Height          =   270
      Left            =   150
      TabIndex        =   7
      Top             =   105
      Width           =   2580
   End
End
Attribute VB_Name = "FrmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(c) 2007-2009 - Diomidisk
Option Explicit

Dim myH As Single, myW As Single
Dim tmpW As Single, tmpH As Single
Dim mColor As OLE_COLOR
Dim mImage As String
Dim Canceled As Boolean
' Display the form. Return True if the user cancels.
Public Function ShowForm(cWidth As Single, _
                         cHeight As Single, _
                         Optional cImage As String = "", _
                         Optional cColor As OLE_COLOR) As Boolean
    
    myW = cWidth
    myH = cHeight
    
    tmpW = cWidth
    tmpH = cHeight
    
    mImage = cImage
    mColor = cColor
    PicColor.BackColor = mColor
    TxtHeight.Text = cHeight
    TxtWidth.Text = cWidth
    
    If gScaleMode = vbPixels Then
       OptionType(0).Value = True
    ElseIf gScaleMode = vbInches Then
       OptionType(1).Value = True
    ElseIf gScaleMode = vbMillimeters Then
       OptionType(2).Value = True
    End If
        
    If gPrintetOrientation = 1 Then
       optOrien(0).Value = True
       imgPrinterOrien.Picture = imgPage(0).Picture
    Else
       optOrien(1).Value = True
       imgPrinterOrien.Picture = imgPage(1).Picture
    End If
    
    ' Display the form.
    Show vbModal
    ShowForm = Canceled
    
    cWidth = myW
    cHeight = myH
   
'    If OptionColor.Value Then
       cColor = PicColor.BackColor
       cImage = ""
'    End If
'    If OptionColor.Value = False Then
'       cColor = vbWhite ' -1
'       cImage =  ""
'    End If
    Unload Me
End Function

Private Sub CmdCancel_Click()
    Canceled = True
    Hide
   
End Sub

Private Sub cmdColor_Click()
        OpenColorDialog PicColor
End Sub

Private Sub ComOK_Click()

    Canceled = False
    Hide
   
End Sub

Private Sub Form_Load()
    
    TxtHeight.Text = myH
    TxtWidth.Text = myW
    tmpW = myW
    tmpH = myH
    'List1.AddItem "Page A0": List1.ListIndex(List1.NewIndex) = 0
    'List1.AddItem "Page A1": List1.ListIndex(List1.NewIndex) = 1
    'List1.AddItem "Page A2": List1.ListIndex(List1.NewIndex) = 2
    'List1.AddItem "Page A3": List1.ListIndex(List1.NewIndex) = 3
    List1.AddItem "Page A4": List1.ItemData(List1.NewIndex) = 4
   ' List1.AddItem "Page A5": List1.ItemData(List1.NewIndex) = 5
   ' List1.AddItem "Page A6": List1.ItemData(List1.NewIndex) = 6
    List1.AddItem "Page 8.5''x11''": List1.ItemData(List1.NewIndex) = 7
    
    List1.ListIndex = 0
End Sub

Private Sub List1_Click()
    
    optOrien(0).Value = True
    
    Select Case List1.ItemData(List1.ListIndex)
    Case 0
        myW = 3179: myH = 4494
        OptionType(2).Value = True
    Case 1
        myW = 2245: myH = 3179
        OptionType(2).Value = True
    Case 2
        myW = 1587: myH = 2245
        OptionType(2).Value = True
    Case 3
        myW = 1123: myH = 1587
        OptionType(2).Value = True
    Case 4
       'myW = 800: myH = 600
       myW = 794: myH = 1123
       OptionType(2).Value = True
    Case 5
       myW = 559: myH = 794
       OptionType(2).Value = True
       'myW = 1024: myH = 768
    Case 6
        myW = 397: myH = 559
       OptionType(2).Value = True
       'myW = 1280: myH = 1024
    Case 7
       myW = 816: myH = 1056
       OptionType(1).Value = True
      ' myW = 85: myH = 54
    End Select
    tmpW = myW
    tmpH = myH
    
    If OptionType(0).Value = True Then
        TxtWidth.Text = myW
        TxtHeight.Text = myH
    ElseIf OptionType(1).Value = True Then
        TxtWidth.Text = Round(ScaleX(myW, vbPixels, vbInches), 2)
        TxtHeight.Text = Round(ScaleY(myH, vbPixels, vbInches), 2)
    ElseIf OptionType(2).Value = True Then
        TxtWidth.Text = Round(ScaleX(myW, vbPixels, vbMillimeters), 0)
        TxtHeight.Text = Round(ScaleY(myH, vbPixels, vbMillimeters), 0)
    End If
    
End Sub


Private Sub OptionType_Click(Index As Integer)
    Select Case Index
    Case 0
        TxtWidth.Text = Round(myW)
        TxtHeight.Text = Round(myH)
        gScaleMode = vbPixels
    Case 1
        TxtWidth.Text = Round(ScaleX(myW, vbPixels, vbInches), 2)
        TxtHeight.Text = Round(ScaleY(myH, vbPixels, vbInches), 2)
        gScaleMode = vbInches
    Case 2
        TxtWidth.Text = Round(ScaleX(myW, vbPixels, vbMillimeters), 0)
        TxtHeight.Text = Round(ScaleY(myH, vbPixels, vbMillimeters), 0)
        gScaleMode = vbMillimeters
    End Select
    
    SaveSetting App.ProductName, "Main", "ScaleMode", Trim(Str(gScaleMode))

End Sub

Private Sub optOrien_Click(Index As Integer)
   
    Select Case Index
    Case 0
          imgPrinterOrien.Picture = imgPage(Index).Picture
           myW = tmpW
           myH = tmpH
           gPrintetOrientation = vbPRORPortrait
    Case 1
          imgPrinterOrien.Picture = imgPage(Index).Picture
           myW = tmpH
           myH = tmpW
           gPrintetOrientation = vbPRORLandscape
    End Select
    
    If OptionType(0).Value = True Then
        TxtWidth.Text = myW
        TxtHeight.Text = myH
    ElseIf OptionType(1).Value = True Then
        TxtWidth.Text = Round(ScaleX(myW, vbPixels, vbInches), 2)
        TxtHeight.Text = Round(ScaleY(myH, vbPixels, vbInches), 2)
    ElseIf OptionType(2).Value = True Then
        TxtWidth.Text = Round(ScaleX(myW, vbPixels, vbMillimeters), 0)
        TxtHeight.Text = Round(ScaleY(myH, vbPixels, vbMillimeters), 0)
    End If
End Sub

Private Sub txtHeight_Change()
On Error Resume Next

    If OptionType(1).Value = True Then
        myH = Round(ScaleY(Format(TxtHeight.Text), vbInches, vbPixels), 2)
    ElseIf OptionType(2).Value = True Then
        myH = Round(ScaleY(Format(TxtHeight.Text), vbMillimeters, vbPixels), 2)
    Else
        myH = TxtHeight.Text
    End If
End Sub

Private Sub txtWidth_Change()
On Error Resume Next

    If OptionType(1).Value = True Then
        myW = Round(ScaleX(Format(TxtWidth.Text), vbInches, vbPixels), 2)
    ElseIf OptionType(2).Value = True Then
        myW = Round(ScaleX(Format(TxtWidth.Text), vbMillimeters, vbPixels), 2)
    Else
        myW = TxtWidth.Text
    End If
End Sub


