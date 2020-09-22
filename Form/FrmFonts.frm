VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmFonts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Text"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlToolBar 
      Left            =   60
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   37
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":0000
            Key             =   "New"
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":0352
            Key             =   "Open"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":06A4
            Key             =   "Save"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":09F6
            Key             =   "Print"
            Object.Tag             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":0D48
            Key             =   "Export"
            Object.Tag             =   "Export"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":109A
            Key             =   "Cut"
            Object.Tag             =   "Cut"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":13EC
            Key             =   "Copy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":173E
            Key             =   "Paste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":1A90
            Key             =   "Undo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":1DE2
            Key             =   "Redo"
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":2134
            Key             =   "Delete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":2486
            Key             =   "TextLeft"
            Object.Tag             =   "TextLeft"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":27D8
            Key             =   "TextCenter"
            Object.Tag             =   "TextCenter"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":2B2A
            Key             =   "TextRight"
            Object.Tag             =   "TextRight"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":2E7C
            Key             =   "Bold"
            Object.Tag             =   "Bold"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":31CE
            Key             =   "Italic"
            Object.Tag             =   "Italic"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":3520
            Key             =   "Underline"
            Object.Tag             =   "Underline"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":3872
            Key             =   "Strikethru"
            Object.Tag             =   "Strikethru"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":3BC4
            Key             =   "SelectAll"
            Object.Tag             =   "SelectAll"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":3F16
            Key             =   "UnselectAll"
            Object.Tag             =   "UnselectAll"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":4268
            Key             =   "AlignLeft"
            Object.Tag             =   "AlignLeft"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":45BA
            Key             =   "AlignCenterVertical"
            Object.Tag             =   "AlignCenterVertical"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":490C
            Key             =   "AlignRight"
            Object.Tag             =   "AlignRight"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":4C5E
            Key             =   "AlignTop"
            Object.Tag             =   "AlignTop"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":4FB0
            Key             =   "AlignCenterHorizontal"
            Object.Tag             =   "AlignCenterHorizontal"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":5302
            Key             =   "AlignBottom"
            Object.Tag             =   "AlignBottom"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":5654
            Key             =   "AlignCenterVerticalHorizontal"
            Object.Tag             =   "AlignCenterVerticalHorizontal"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":59A6
            Key             =   "BringToFront"
            Object.Tag             =   "BringToFront"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":5CF8
            Key             =   "SendToBack"
            Object.Tag             =   "SendToBack"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":604A
            Key             =   "BringForward"
            Object.Tag             =   "BringForward"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":639C
            Key             =   "SendBackward"
            Object.Tag             =   "SendBackward"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":66EE
            Key             =   "Group"
            Object.Tag             =   "Group"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":6A40
            Key             =   "Ungroup"
            Object.Tag             =   "Ungroup"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":6D92
            Key             =   "Zoom100"
            Object.Tag             =   "Zoom100"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":70E4
            Key             =   "Zoom-"
            Object.Tag             =   "Zoom-"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":7436
            Key             =   "Zoom+"
            Object.Tag             =   "Zoom+"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFonts.frx":7788
            Key             =   "ZoomAll"
            Object.Tag             =   "ZoomAll"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   3450
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   420
      Width           =   6525
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2250
      TabIndex        =   1
      Top             =   4035
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3315
      TabIndex        =   0
      Top             =   4035
      Width           =   1000
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imlToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "TextLeft"
            Object.ToolTipText     =   "Align Text Left"
            Object.Tag             =   "AlignText"
            ImageKey        =   "TextLeft"
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "TextCenter"
            Object.ToolTipText     =   "Align Text Center"
            Object.Tag             =   "AlignText"
            ImageKey        =   "TextCenter"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "TextRight"
            Object.ToolTipText     =   "Align Text Right"
            Object.Tag             =   "AlignText"
            ImageKey        =   "TextRight"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            Object.Tag             =   "Bold"
            ImageKey        =   "Bold"
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            Object.Tag             =   "Italic"
            ImageKey        =   "Italic"
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            Object.Tag             =   "Underline"
            ImageKey        =   "Underline"
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Strikethru"
            Object.ToolTipText     =   "Strikethru"
            Object.Tag             =   "Strikethru"
            ImageKey        =   "Strikethru"
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.ComboBox CboFontSize 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "FrmFonts.frx":7ADA
         Left            =   3540
         List            =   "FrmFonts.frx":7ADC
         TabIndex        =   5
         Text            =   "15"
         ToolTipText     =   "Font Size"
         Top             =   15
         Width           =   705
      End
      Begin VB.ComboBox CboFontName 
         Height          =   315
         ItemData        =   "FrmFonts.frx":7ADE
         Left            =   1500
         List            =   "FrmFonts.frx":7AE0
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "CboFontName"
         ToolTipText     =   "Font Name"
         Top             =   15
         Width           =   1905
      End
   End
End
Attribute VB_Name = "FrmFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(c) 2007 diomidisk
Option Explicit

Private Canceled As Boolean
Private sFonts As New StdFont

' Display the form. Return True if the user cancels.
Public Function ShowForm(ByRef nfonts As StdFont, ByRef txt As String, ByRef mAlign As Integer) As Boolean
Dim t As Integer, IndexS As Integer, Idx As Integer
    ' Assume we will cancel.
    Canceled = True
    
    nfonts.Size = Round(nfonts.Size, 0)
    
    CboFontSize.Clear
    Idx = 0
    For t = 5 To 12
        CboFontSize.AddItem t
        Idx = Idx + 1
        If t = Round(nfonts.Size, 0) Then IndexS = Idx
    Next
    For t = 14 To 28 Step 2
        CboFontSize.AddItem t
        Idx = Idx + 1
        If t = Round(nfonts.Size, 0) Then IndexS = Idx
    Next
    CboFontSize.AddItem 36
    If 36 = Round(nfonts.Size, 0) Then IndexS = Idx + 1
    CboFontSize.AddItem 48
    If 48 = Round(nfonts.Size, 0) Then IndexS = Idx + 2
    CboFontSize.AddItem 72
    If 72 = Round(nfonts.Size, 0) Then IndexS = Idx + 3
    CboFontSize.AddItem 100
    If 100 = Round(nfonts.Size, 0) Then IndexS = Idx + 4
    CboFontSize.AddItem 200
    If 200 = Round(nfonts.Size, 0) Then IndexS = Idx + 5
    CboFontSize.AddItem 300
    If 300 = Round(nfonts.Size, 0) Then IndexS = Idx + 6

    If IndexS <= CboFontSize.ListCount Then
        CboFontSize.ListIndex = IndexS - 1
    End If
    CboFontSize.Text = Round(nfonts.Size, 0)
        
    If CboFontName.ListCount = 0 Then
         Screen.MousePointer = 11
         LoadFonts CboFontName
        CboFontName.Text = "Arial"
        Screen.MousePointer = 0
    End If
    
    'Text1.Font.Bold = nFonts.Bold
    Text1.Text = txt
    Text1.Font = nfonts
    Text1.Font.Charset = nfonts.Charset
    Text1.Font.Italic = nfonts.Italic
    If nfonts.Name <> "" Then CboFontName.Text = nfonts.Name
    Text1.Font.Name = nfonts.Name
    Text1.Font.Size = nfonts.Size
    Text1.Font.Strikethrough = nfonts.Strikethrough
    Text1.Font.Underline = nfonts.Underline
    Text1.Font.Weight = nfonts.Weight
    
    ' Display the form.
    Show vbModal

    ShowForm = Canceled
    
    If Not Canceled Then
        On Error Resume Next
        Text1.Font = sFonts
        nfonts.Bold = Text1.Font.Bold
        nfonts.Charset = Text1.Font.Charset
        nfonts.Italic = Text1.Font.Italic
        nfonts.Name = Text1.Font.Name
        nfonts.Size = Text1.Font.Size
        nfonts.Strikethrough = Text1.Font.Strikethrough
        nfonts.Underline = Text1.Font.Underline
        nfonts.Weight = Text1.Font.Weight
        mAlign = Text1.Alignment
        txt = Text1.Text
        On Error GoTo 0
    End If
    Unload Me
End Function


Private Sub CboFontName_Change()
       CboFontName_Click
End Sub

Private Sub CboFontName_Click()
     sFonts.Italic = False
          
     Text1.Font.Name = CboFontName.Text
     sFonts.Name = CboFontName.Text
     Text1.Font.Bold = sFonts.Bold
     Text1.Font.Charset = sFonts.Charset
     Text1.Font.Italic = sFonts.Italic
     Text1.Font.Size = sFonts.Size
     Text1.Font.Strikethrough = sFonts.Strikethrough
     Text1.Font.Underline = sFonts.Underline
     Text1.Font.Weight = sFonts.Weight
     
End Sub

Private Sub CboFontSize_Click()
     If Val(CboFontSize.Text) > 0 Then
         Text1.FontSize = Val(CboFontSize.Text)
         sFonts.Size = Val(CboFontSize.Text)
     End If
End Sub

Private Sub CmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()
    Canceled = False
    Hide
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
       Text1.Alignment = 0
    Case 2
       Text1.Alignment = 2
    Case 3
       Text1.Alignment = 1
    Case 4 '-
    Case 5
       Text1.FontBold = Button.Value
    Case 6
       Text1.FontItalic = Button.Value
    Case 7
       Text1.FontUnderline = Button.Value
    Case 8
       Text1.FontStrikethru = Button.Value
    End Select
End Sub


