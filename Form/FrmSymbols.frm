VERSION 5.00
Begin VB.Form FrmSymbols 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Symbols"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ArtDraw.CtrFonts CtrFonts1 
      Height          =   6630
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   11695
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   24
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   6690
      Width           =   2775
   End
End
Attribute VB_Name = "FrmSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(c) 2007 Diomidisk

Private Sub Command1_Click()
   Dim PointCoods() As PointAPI, PointType() As Byte, iCounter As Long, i As Long
    
    CtrFonts1.ExportSymbol iCounter
    
    If iCounter > 0 Then
       ReDim PointCoods(iCounter)
       ReDim PointType(iCounter)
       For i = 0 To iCounter
          CtrFonts1.GetPointCoods i, PointCoods(i).X, PointCoods(i).Y
          PointType(i) = CtrFonts1.GetPointType(i)
       Next
       FormatData PointCoods, PointType, iCounter - 1
    End If
End Sub

Private Sub Form_Load()
   Me.Move Screen.Width - Me.Width, (Screen.Height - Me.Height) / 2
      FormOnTop Me, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
      m_FormSymbolView = False
      FormOnTop Me, False
End Sub

Private Sub FormatData(ByRef PointCoords() As PointAPI, ByRef PointType() As Byte, ByVal iCounter As Long)
    Dim txt As String, OldTxt As String, i As Long

    txt = txt & " DrawWidth(1)"
    txt = txt & " DrawStyle(0)"
    txt = txt & " ForeColor(0)"
    txt = txt & " FillColor(0)"
    txt = txt & " FillStyle(0)"
    
    txt = txt & " TextDraw()"
    txt = txt & " TypeDraw(" + Format$(dPolydraw) + ")"
    txt = txt & " CurrentX(0)"
    txt = txt & " CurrentY(0)"
    txt = txt & " TypeFill(0)"
    txt = txt & " Pattern()"
    txt = txt & " Shade(False)"
    
    txt = txt & " Bold(0)"
    txt = txt & " Charset(0)"
    txt = txt & " Italic(0)"
    txt = txt & " Name()"
    txt = txt & " Size(0)"
    txt = txt & " Strikethrough(0)"
    txt = txt & " Underline(0)"
    txt = txt & " Weight(400)"
    txt = txt & " Angle(0)"
    
    txt = txt & vbCr & "Transformation(1 0 0 0 1 0 0 0 1 )"
    txt = txt & " IsClosed(True)"
    txt = txt & " NumPoints(" & Format$(iCounter + 1) & ")"
    
    For i = 0 To iCounter
            txt = txt & vbCrLf & "    X(" & Format$(PointCoords(i).X) & ")"
            txt = txt & " Y(" & Format$(PointCoords(i).Y) & ")"
            txt = txt & " P(" & Format$(PointType(i)) & ")"
    Next i

    txt = "PolyDraw(PolyDraw(" & txt & ")"
    
    OldTxt = Clipboard.GetText
   
    Clipboard.SetText txt
    frmVbDraw.DrawControl1.PasteObject
    Clipboard.SetText OldTxt
    
End Sub


