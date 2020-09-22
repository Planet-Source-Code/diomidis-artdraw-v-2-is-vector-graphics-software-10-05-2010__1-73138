VERSION 5.00
Begin VB.UserControl ColorPalette 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   55
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ColorPalette.ctx":0000
   Begin VB.CommandButton CommandOpen 
      Height          =   270
      Left            =   3030
      Picture         =   "ColorPalette.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Select Palette"
      Top             =   15
      Width           =   270
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   10
      Left            =   0
      Max             =   255
      SmallChange     =   10
      TabIndex        =   0
      Top             =   270
      Width           =   3285
   End
   Begin VB.PictureBox PicturePalette 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   0
      MouseIcon       =   "ColorPalette.ctx":05C4
      MousePointer    =   99  'Custom
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   1
      ToolTipText     =   "Left click Fill color - Right click Border color"
      Top             =   0
      Width           =   2985
   End
End
Attribute VB_Name = "ColorPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(c) 2007 - Diomidisk

'Default Property Values:
Const m_def_FileNamePalette = ""

'Property Variables:
Dim m_FileNamePalette As String

Dim m_ColorList() As Long
Dim Maxm_Col As Integer
Dim mH As Long
Dim mView As Integer
Dim m_ColorFill As Long
Dim m_IdFill As Integer
Dim m_ColorBorder As Long
Dim m_IdBorder As Integer

'Event Declarations:
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event ColorSelected(Button As Integer, cm_Color As Long)
Event ColorOver(cm_Color As Long)

Private Sub CommandOpen_Click()
Dim user_canceled As Boolean
Dim FileNamePalette As String

    user_canceled = FrmPalette.ShowForm(FileNamePalette)
    Unload FrmPalette

    ' If the user canceled, do no more.
    If user_canceled Then Exit Sub
 
    LoadPalette FileNamePalette
    PicturePalette.SetFocus
    
End Sub

Private Sub HScroll1_Change()
       mView = HScroll1.Value
       LoadPalette
End Sub


Private Sub UserControl_Resize()
      
      mH = 18
      PicturePalette.Move 0, 0, 256 * mH + 18, mH - 1
      HScroll1.Move 0, 18, UserControl.ScaleWidth
      CommandOpen.Move UserControl.ScaleWidth - 18, 0, 18, 18
     
      LoadPalette
End Sub

Public Sub LoadPalette(Optional PalFile As String)
   
On Error Resume Next
Dim ff As Integer
Dim m_Str As String
Dim n As Integer
Dim m_Qty As Integer
Dim m_Row As Integer
Dim m_Col As Integer

   If PalFile <> "" Then FileNamePalette = PalFile

ff = FreeFile

If PalFile = "" Or Dir(PalFile) = "" Then
    If FileNamePalette <> "" Then
      If FileExists(FileNamePalette) Then
         PalFile = FileNamePalette
      Else
        PalFile = App.Path & "\Default.pal"
      End If
   Else
      PalFile = App.Path & "\Default.pal"
   End If
Else
    PalFile = App.Path & "\Default.pal"
End If
  
If FileExists(App.Path + "\Palette\Default.pal") = True And FileExists(App.Path & "\Default.pal") Then
   FileCopy App.Path + "\Palette\Default.pal", App.Path & "\Default.pal"
   Kill App.Path + "\Palette\Default.pal"
End If
If FileExists(PalFile) = False Then
   FileCopy App.Path + "\Palette\ArtDraw.pal", App.Path & "\Default.pal"
End If
If Dir(PalFile) <> "" Then
Open PalFile For Input As #ff
    Input #ff, m_Str 'JASC-PAL
    If UCase(m_Str) <> "JASC-PAL" Then
       Close #ff
        Exit Sub
    End If
    Input #ff, m_Str '0010
    Input #ff, m_Str '256 (m_Color qty)
    m_Qty = Int(m_Str)
    
    ReDim m_ColorList(Int(m_Qty))
    n = 0
    While Not EOF(ff)
       Input #ff, m_Str
ragain:
       m_Str = Replace(m_Str, "  ", " ")
       If InStr(1, m_Str, "  ") Then GoTo ragain
       m_ColorList(n) = RGB(Split(m_Str, " ")(0), Split(m_Str, " ")(1), Split(m_Str, " ")(2))
       n = n + 1
    Wend
Close #ff

   HScroll1.Max = m_Qty - (UserControl.ScaleWidth \ mH) + 1
   PicturePalette.Line (0, 0)-(PicturePalette.Width, PicturePalette.Height), QBColor(15), BF
   PicturePalette.Line (0, 0)-(18, 18), QBColor(0)
   PicturePalette.Line (0, 18)-(18, 0), QBColor(0)
     
   Maxm_Col = m_Qty
   Fillm_ColorPalette
    
End If

Exit Sub

ErrLoad:
   Close #ff
End Sub

Sub Fillm_ColorPalette()
 m_Col = 1
    For n = mView To Maxm_Col - 1
       PicturePalette.Line (m_Col * mH, 0)-Step(mH, mH), m_ColorList(n), BF
       PicturePalette.Line (m_Col * mH, 0)-Step(mH, mH), , B
       m_Col = m_Col + 1
    Next n
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_FileNamePalette = m_def_FileNamePalette
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_FileNamePalette = PropBag.ReadProperty("FileNamePalette", m_def_FileNamePalette)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    m_ColorFill = RGB(255, 255, 255)
    m_ColorBorder = RGB(0, 0, 0)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("FileNamePalette", m_FileNamePalette, m_def_FileNamePalette)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
End Sub

Private Sub PicturePalette_Click()
    RaiseEvent Click
End Sub

Private Sub PicturePalette_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub PicturePalette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim tm_Color As Long
Dim tInd As Integer
   If (X \ mH) - 1 = -1 Then
       tInd = -1
    Else
       tInd = (mView + X \ mH) - 1
    End If
    ''Debug.Print tInd
    If tInd > UBound(m_ColorList) Then Exit Sub
    If tInd >= 0 Then
        tm_Color = m_ColorList(tInd)
    End If
    If tInd = -1 Then tm_Color = -1
       
  If tm_Color <> -1 Then
     If Button = 1 Then
        m_IdFill = tInd
        m_ColorFill = tm_Color
     Else
        m_IdBorder = tInd
        m_ColorBorder = tm_Color
     End If
  End If
  RaiseEvent ColorSelected(Button, tm_Color)
  RaiseEvent MouseDown(Button, Shift, X, Y)
    
End Sub

Private Sub PicturePalette_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

Dim tm_Color As Long
Dim tInd As Integer

    tInd = (mView + X \ mH) - 1
    If tInd > UBound(m_ColorList) Or tInd = -1 Then Exit Sub
    tm_Color = m_ColorList(tInd)
    If tInd = -1 Then tm_Color = -1
    RaiseEvent ColorOver(tm_Color)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub PicturePalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FileNamePalette() As String
    FileNamePalette = m_FileNamePalette
End Property

Public Property Let FileNamePalette(ByVal New_FileNamePalette As String)
    m_FileNamePalette = New_FileNamePalette
    PropertyChanged "FileNamePalette"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

