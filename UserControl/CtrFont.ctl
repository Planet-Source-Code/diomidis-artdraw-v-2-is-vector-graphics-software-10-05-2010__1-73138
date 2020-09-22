VERSION 5.00
Begin VB.UserControl CtrFonts 
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   ScaleHeight     =   6705
   ScaleWidth      =   3030
   ToolboxBitmap   =   "CtrFont.ctx":0000
   Begin VB.TextBox TextSize 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1980
      TabIndex        =   7
      Top             =   6330
      Width           =   645
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   285
      Left            =   2640
      Max             =   300
      Min             =   8
      TabIndex        =   6
      Top             =   6330
      Value           =   20
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1755
      Left            =   2670
      Max             =   300
      Min             =   20
      TabIndex        =   4
      Top             =   600
      Value           =   20
      Width           =   240
   End
   Begin VB.TextBox txtSymbol 
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      ToolTipText     =   "Select symbol"
      Top             =   2415
      Width           =   2850
   End
   Begin VB.ComboBox cboFont 
      Height          =   315
      ItemData        =   "CtrFont.ctx":0312
      Left            =   60
      List            =   "CtrFont.ctx":0322
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Symbols"
      Top             =   150
      Width           =   2835
   End
   Begin VB.PictureBox Picturefont 
      BackColor       =   &H00FFFFFF&
      Height          =   1755
      Left            =   45
      ScaleHeight     =   1695
      ScaleWidth      =   2550
      TabIndex        =   1
      ToolTipText     =   "Preview select"
      Top             =   585
      Width           =   2610
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   2310
      Pattern         =   "*.ttf;*.otf"
      TabIndex        =   0
      Top             =   930
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LabelInfo 
      Caption         =   "Point :"
      Height          =   255
      Left            =   105
      TabIndex        =   5
      Top             =   6345
      Width           =   1440
   End
End
Attribute VB_Name = "CtrFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(c) 2007 - Diomidisk

Private m_Preview As CFontPreview
Private FontsfullName() As String
Private FontName As String
'Default Property Values:
Const m_def_Symbol = ""
'Property Variables:
Dim m_Symbol As String

Dim PointCoods() As PointAPI, PointType() As Byte

Private Sub ComboFonts_Click()

    txtSymbol.Font = cboFont.Text
    txtSymbol.Font.Size = 20
    txtSymbol.Text = ""
    i = 33
    For Row = 1 To 45
       For Col = 1 To 5
          txtSymbol.Text = txtSymbol.Text + Trim(Chr(i))
          i = i + 1
          If i > 255 Then GoTo endfill
       Next
    Next
endfill:
   txtSymbol.Text = Replace(txtSymbol.Text, " ", "")
End Sub

Private Sub cboFont_Click()
     If Not m_Preview Is Nothing Then
         Set m_Preview = Nothing
     End If
      
     If m_Preview Is Nothing Then
        NewPreview
     End If
    If cboFont.Text = "Webdings" Or cboFont.Text = "Wingdings" Or cboFont.Text = "Wingdings 2" Or cboFont.Text = "Wingdings 3" Then
       FontName = cboFont.Text
    Else
       If Len(FontsfullName(cboFont.ListIndex)) > 0 Then
         m_Preview.FontFile = FontsfullName(cboFont.ListIndex)
          FontName = m_Preview.FaceName
       Else
         FontName = ""
       End If
    End If
    If Len(FontName) > 0 Then
       txtSymbol.Font = FontName
       txtSymbol.SelStart = 1
       txtSymbol.FontSize = 24
       txtSymbol.FontStrikethru = False
       txtSymbol.FontUnderline = False
       txtSymbol.FontBold = False
       txtSymbol.FontItalic = False
       Picturefont.Picture = LoadPicture()
       Picturefont.Font = txtSymbol.Font
       Picturefont.FontStrikethru = txtSymbol.FontStrikethru
       Picturefont.FontUnderline = txtSymbol.FontUnderline
       Picturefont.FontBold = txtSymbol.FontBold
       Picturefont.FontItalic = txtSymbol.FontItalic
    End If
End Sub

Private Sub InsertFont_Click()
      
End Sub

Private Sub Picturefont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Picturefont.ToolTipText = "Preview symbol size=" + Str(Picturefont.FontSize)
End Sub

Private Sub txtSymbol_Click()
  Dim iCounter As Long
  If m_Preview Is Nothing Then Exit Sub
    If Len(FontName) > 0 Then
       txtSymbol.SelLength = 1
       Picturefont.Picture = LoadPicture()
       Picturefont.Font = FontName
       Picturefont.FontStrikethru = txtSymbol.FontStrikethru
       Picturefont.FontUnderline = txtSymbol.FontUnderline
       Picturefont.FontBold = txtSymbol.FontBold
       Picturefont.FontItalic = txtSymbol.FontItalic
       Picturefont.Cls
       Picturefont.CurrentX = 0
       Picturefont.CurrentY = 0
       If Picturefont.Font.Size > 100 Then Picturefont.Font.Size = 100
       If txtSymbol.SelStart + 33 > 255 Then
          Picturefont.Print Chr(txtSymbol.SelStart + 32)
       Else
          Symbol = Chr(txtSymbol.SelStart + 33)
          Picturefont.Print Chr(txtSymbol.SelStart + 33)
          ExportSymbol iCounter
          LabelInfo.Caption = "Point :" + Str(iCounter)
       End If
   End If
End Sub

Private Sub txtSymbol_KeyUp(KeyCode As Integer, Shift As Integer)
      If KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Then
        txtSymbol_Click
      End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim i As Long
    
    Set txtSymbol.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Symbol = PropBag.ReadProperty("Symbol", m_def_Symbol)
    
    File1.Path = App.Path + "\fonts"
    ReDim FontsfullName(File1.ListCount - 1 + 4)
    
    For i = 0 To File1.ListCount - 1
    Set m_Preview = New CFontPreview
       File1.ListIndex = i
        m_Preview.FontFile = File1.Path & "\" & File1.Filename
       If Len(m_Preview.FullName) > 0 Then
         FontsfullName(i + 4) = File1.Path & "\" & File1.Filename
         cboFont.AddItem m_Preview.FullName
       End If
       Set m_Preview = Nothing
    Next
    Set m_Preview = New CFontPreview
    
'    If cboFont.ListCount > 0 Then
'       cboFont.ListIndex = 0
'    End If
    
    txtSymbol.FontSize = 24
    For i = 33 To 255
        txtSymbol.Text = txtSymbol.Text + Chr(i)
    Next i
    Picturefont.FontSize = 36
    VScroll1.Value = 36
    If cboFont.ListCount > 0 Then
       cboFont.ListIndex = 0
    End If
End Sub

Private Sub UserControl_Terminate()
   Set m_Preview = Nothing
End Sub

Private Sub VScroll1_Change()
    Picturefont.FontSize = VScroll1.Value
    VScroll2.Value = VScroll1.Value
    txtSymbol_Click
End Sub

Sub NewPreview()
   Set m_Preview = New CFontPreview
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtSymbol.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtSymbol.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get Symbol() As String
    Symbol = m_Symbol
End Property

Public Property Let Symbol(ByVal New_Symbol As String)
    m_Symbol = New_Symbol
    PropertyChanged "Symbol"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Symbol = m_def_Symbol
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Font", txtSymbol.Font, Ambient.Font)
    Call PropBag.WriteProperty("Symbol", m_Symbol, m_def_Symbol)
End Sub

Public Sub ExportSymbol(iCounter As Long)
   Dim nfonts As New StdFont
    If Symbol = "" Then Exit Sub
    
           nfonts.Bold = txtSymbol.Font.Bold
           nfonts.Charset = 0
           nfonts.Italic = False
           nfonts.Name = txtSymbol.Font.Name
           nfonts.Size = 100
           nfonts.Strikethrough = False
           nfonts.Underline = False
           With Picturefont
              .Font.Size = nfonts.Size 'VScroll2.Value
            End With
            Erase PointCoods
            Erase PointType
            iCounter = 0
            'ReadPathText Picturefont, Symbol, PointCoods(), PointType(), iCounter
            
            Call BeginPath(Picturefont.hDC)
                Picturefont.CurrentX = 0
                Picturefont.CurrentY = 0
                Picturefont.Print Symbol
            Call EndPath(Picturefont.hDC)
            iCounter = 0
            iCounter = GetPathAPI(Picturefont.hDC, ByVal 0&, ByVal 0&, 0)

            If (iCounter) Then
                ReDim PointCoods(iCounter - 1)
                ReDim PointType(iCounter - 1)
                'Get the path data from the DC
                Call GetPathAPI(Picturefont.hDC, PointCoods(0), PointType(0), iCounter)
            End If
            
            With Picturefont
              .Font.Size = VScroll2.Value
            End With
End Sub

Public Sub GetPointCoods(iCounter As Long, X As Long, Y As Long)
        If iCounter <= UBound(PointCoods) Then
          X = PointCoods(iCounter).X
          Y = PointCoods(iCounter).Y
        End If
End Sub

Public Function GetPointType(iCounter As Long) As Byte
        If iCounter <= UBound(PointType) Then
           GetPointType = PointType(iCounter)
        End If
End Function

Private Sub VScroll2_Change()
       TextSize.Text = Format$(VScroll2.Value)
End Sub
