VERSION 5.00
Begin VB.Form FrmPalette 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Palette"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdDefault 
      Caption         =   "Default"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5655
      TabIndex        =   8
      Top             =   2415
      Width           =   1200
   End
   Begin VB.CommandButton CmdSavePalette 
      Caption         =   "Save Palette"
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   1875
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000013&
      Height          =   3495
      Left            =   3345
      MouseIcon       =   "FrmPalette.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   5
      Top             =   390
      Width           =   2025
      Begin VB.PictureBox PictureColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   0
         Left            =   15
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   6
         Top             =   60
         Width           =   120
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   870
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   150
      Pattern         =   "*.pal"
      TabIndex        =   1
      Top             =   390
      Width           =   2940
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Color palette "
      Height          =   210
      Left            =   3345
      TabIndex        =   3
      Top             =   135
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Select palette :"
      Height          =   210
      Left            =   165
      TabIndex        =   2
      Top             =   105
      Width           =   2850
   End
End
Attribute VB_Name = "FrmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(c) 2007 Diomidisk

Option Explicit
Private Canceled As Boolean
Private SelectColor As Boolean

Dim ColorList() As Long
Const MaxCol = 12
Const TSize = 10

Private Sub CmdCancel_Click()
    Me.Hide
End Sub

Private Sub CmdDefault_Click()
     If MsgBox("You want the new palette to make default?", vbInformation + vbYesNo, "Save Palette") = vbYes Then
        
        If FileExists(File1.Path + "\" + File1.Filename) Then
           FileCopy File1.Path + "\" + File1.Filename, App.Path + "\Palette\tmp.pal"
           If FileExists(App.Path + "\Default.pal") Then Kill App.Path + "\Default.pal"
           FileCopy App.Path + "\Palette\tmp.pal", App.Path + "\Default.pal"
           Kill App.Path + "\Palette\tmp.pal"
           File1.Refresh
           'CmdSavePalette.Enabled = False
           CmdDefault.Enabled = False
        End If
        
     End If
End Sub

Private Sub cmdOK_Click()
    Canceled = False
    Me.Hide
End Sub

Private Sub CmdSavePalette_Click()
       SavePalette
       'CmdSavePalette.Enabled = False
       CmdDefault.Enabled = False
End Sub

Private Sub File1_Click()
   If FileExists(File1.Path + "\" + File1.Filename) Then
      Me.MousePointer = 11
      LoadPalette File1.Path + "\" + File1.Filename
      Me.MousePointer = 0
      'CmdSavePalette.Enabled = False
      CmdDefault.Enabled = True
      SelectColor = False
   End If
End Sub

Private Sub File1_KeyUp(KeyCode As Integer, Shift As Integer)
     
     Dim id As Long
     id = File1.ListIndex
     If KeyCode = 46 Then
        Kill File1.Path + "\" + File1.Filename
        File1.Refresh
        If File1.ListIndex = -1 Then Exit Sub
        File1.ListIndex = id
     End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    File1.Path = App.Path + "\Palette"
    File1.Refresh
    For i = 1 To 256
        Load PictureColor(i)
    Next
    If File1.ListCount > 0 Then File1.ListIndex = 0
End Sub

' Display the form. Return True if the user cancels.
Public Function ShowForm(FileNamePalette As String) As Boolean
    ' Assume we will cancel.
    Canceled = True

    ' Display the form.
    Show vbModal

    ShowForm = Canceled
    If Not Canceled Then
        On Error Resume Next
        FileNamePalette = App.Path + "\Degault.pal" 'File1.Path + "\" + File1.Filename
        On Error GoTo 0
    End If
End Function

Public Sub LoadPalette(Optional PalFile As String)
On Error Resume Next
Dim ff As Integer
Dim tStr As String
Dim n As Integer
Dim cQty As Integer
Dim Row As Integer
Dim Col As Integer, i As Integer

    ff = FreeFile

    If PalFile = "" Or Dir(PalFile) = "" Then PalFile = App.Path & "\Default.pal"

    If Dir(PalFile) <> "" Then
        Open PalFile For Input As #ff
            Input #ff, tStr$ 'JASC-PAL
            If UCase(tStr) <> "JASC-PAL" Then
                Close #ff
            Exit Sub
            End If
        Input #ff, tStr$ '0010
        Input #ff, tStr$ '256 (color qty)
        cQty = Int(tStr)
        ReDim ColorList(Int(cQty))
    n = 0
    
    While Not EOF(ff)
        Input #ff, tStr$
        ColorList(n) = RGB(Val(Split(tStr, " ")(0)), Val(Split(tStr, " ")(1)), Val(Split(tStr, " ")(2)))
        n = n + 1
    Wend
 Close #ff
 
 Col = 0
 Row = 0
 
 Picture1.ScaleWidth = 12 * 10
 Picture1.Enabled = False
    For n = 0 To cQty - 1
      ' Picture1.Line (Col * TSize, Row * TSize)-(Col * TSize + TSize, Row * TSize + TSize), ColorList(n), BF
       PictureColor(n).Move Col * TSize, Row * TSize, TSize, TSize
       PictureColor(n).BackColor = ColorList(n)
       PictureColor(n).Visible = True
       PictureColor(n).Enabled = True
       Col = Col + 1
       If Col = MaxCol Then
          Col = 0
          Row = Row + 1
        End If
    Next n
    For i = n + 1 To 256 'cQty - 1
       PictureColor(n).Visible = False
       PictureColor(n).Enabled = False
    Next
 Picture1.Enabled = True
 Picture1.ScaleMode = 3
End If

Exit Sub
ErrLoad:
Close #ff
End Sub

Private Sub PictureColor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     Dim NewColor As Long
     Dim C As SelectedColor
     On Error GoTo pm1
     'If File1.ListIndex = -1 Then Exit Sub
     Screen.MousePointer = 11
     NewColor = ColorList(Index)
        
    ColorDialog.rgbResult = NewColor
    C = CommonDialog.ShowColor(Me.hWnd, False)
    
    If C.bCanceled = False Then
       NewColor = C.oSelectedColor
       PictureColor(Index).BackColor = NewColor
       ColorList(Index) = NewColor
       SelectColor = True
      CmdSavePalette.Enabled = True
    End If
    
pm1:
     Screen.MousePointer = 0
     On Error GoTo 0
End Sub

Sub SavePalette()
    Dim PalFile As String, PalFile1 As String, ff As Long, SavePal As Boolean
    Dim IR As Long, IG As Long, IB As Long, i As Long
    If File1.ListIndex <> -1 Then
      If FileExists(File1.Path + "\" + File1.Filename) Then
         PalFile1 = File1.Path + "\" + File1.Filename
      End If
    End If
    PalFile = App.Path & "\NewPalette.pal"
             
    ff = FreeFile
    Open PalFile For Output As #ff
       Print #ff, "JASC-PAL"
       Print #ff, "0100"
       Print #ff, Trim(Str(UBound(ColorList)))
       For i = 0 To UBound(ColorList)
          SplitRGB ColorList(i), IR, IG, IB
          Print #ff, Trim(Str(IR)) + Str(IG) + Str(IB)
       Next
    Close #ff
   
    If PalFile1 <> "" Then
       If FileExists(PalFile1) Then
          If MsgBox("Replace palette '" + PalFile1 + "'?", vbInformation + vbYesNo, "Save Palette") = vbYes Then
             If FileExists(PalFile1) Then
                Kill PalFile1
                FileCopy PalFile, PalFile1
                SavePal = True
             End If
          Else
             SavePal = False
          End If
       End If
    End If
    
    If SavePal = False Then
        Dim sSave As SelectedFile
        If PathExists(App.Path + "\Palette") = False Then MkDir App.Path + "\Palette"
        FileDialog.sFilter = "Palette (*.pal)" & Chr$(0) & "*.pal"
        FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFS_FILE_SAVE_FLAGS
        FileDialog.sInitDir = App.Path + "\Palette"
        FileDialog.sDefFileExt = "*.pal"
        FileDialog.sFile = "NewPalette.pal" + Chr(0)
        sSave = ShowSave(Me.hWnd)
        If Err.Number <> 32755 And sSave.bCanceled = False Then
           Screen.MousePointer = 11
           FileCopy PalFile, sSave.sFile
           Screen.MousePointer = 0
        End If
        FileDialog.sFileTitle = ""
        FileDialog.sInitDir = ""
        FileDialog.sDefFileExt = ""
        FileDialog.flags = 0
    End If

    If FileExists(App.Path & "\NewPalette.pal") Then Kill App.Path & "\NewPalette.pal"
    File1.Refresh
End Sub
