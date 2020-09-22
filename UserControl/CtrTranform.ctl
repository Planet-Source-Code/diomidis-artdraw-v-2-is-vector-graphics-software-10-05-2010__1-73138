VERSION 5.00
Begin VB.UserControl CtrTranform 
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10260
   ScaleHeight     =   3060
   ScaleWidth      =   10260
   ToolboxBitmap   =   "CtrTranform.ctx":0000
   Begin VB.OptionButton Option1 
      Height          =   345
      Index           =   4
      Left            =   1515
      Picture         =   "CtrTranform.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Mirror"
      Top             =   60
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   345
      Index           =   3
      Left            =   1155
      Picture         =   "CtrTranform.ctx":0624
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Skew"
      Top             =   60
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   345
      Index           =   2
      Left            =   795
      Picture         =   "CtrTranform.ctx":0936
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Scale"
      Top             =   60
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2205
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   45
      Width           =   1770
   End
   Begin VB.OptionButton Option1 
      Height          =   345
      Index           =   0
      Left            =   105
      Picture         =   "CtrTranform.ctx":0C48
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Position"
      Top             =   60
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   345
      Index           =   1
      Left            =   480
      Picture         =   "CtrTranform.ctx":0F5A
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Rotation"
      Top             =   60
      Width           =   330
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1700
      Left            =   5625
      ScaleHeight     =   1695
      ScaleWidth      =   1905
      TabIndex        =   16
      Top             =   495
      Visible         =   0   'False
      Width           =   1900
      Begin VB.TextBox TextHSkew 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   300
         TabIndex        =   31
         Text            =   "0"
         Top             =   330
         Width           =   690
      End
      Begin VB.TextBox TextVSkew 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   300
         TabIndex        =   30
         Text            =   "0"
         Top             =   750
         Width           =   690
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   330
         Left            =   990
         Max             =   1800
         Min             =   -1800
         TabIndex        =   29
         Top             =   330
         Width           =   255
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   330
         Left            =   990
         Max             =   2000
         Min             =   -2000
         TabIndex        =   28
         Top             =   765
         Width           =   255
      End
      Begin VB.CommandButton CommandSkew 
         Caption         =   "Apply"
         Height          =   315
         Left            =   200
         TabIndex        =   17
         Top             =   1260
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "deg"
         Height          =   270
         Index           =   8
         Left            =   1365
         TabIndex        =   36
         Top             =   420
         Width           =   300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "deg"
         Height          =   270
         Index           =   7
         Left            =   1365
         TabIndex        =   35
         Top             =   810
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "H :"
         Height          =   225
         Index           =   6
         Left            =   45
         TabIndex        =   34
         Top             =   345
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "V :"
         Height          =   195
         Index           =   5
         Left            =   45
         TabIndex        =   33
         Top             =   780
         Width           =   240
      End
      Begin VB.Label LabelSkew 
         BackStyle       =   0  'Transparent
         Caption         =   "Skew :"
         Height          =   240
         Left            =   30
         TabIndex        =   32
         Top             =   45
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   1700
      Left            =   7530
      ScaleHeight     =   1695
      ScaleWidth      =   1905
      TabIndex        =   12
      Top             =   495
      Visible         =   0   'False
      Width           =   1900
      Begin VB.CheckBox CheckVReflect 
         Height          =   405
         Left            =   315
         Picture         =   "CtrTranform.ctx":11D8
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   705
         Width           =   1305
      End
      Begin VB.CheckBox CheckHReflect 
         Height          =   405
         Left            =   315
         Picture         =   "CtrTranform.ctx":17CA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   315
         Width           =   1305
      End
      Begin VB.CommandButton CommandReflect 
         Caption         =   "Apply"
         Height          =   315
         Left            =   200
         TabIndex        =   13
         Top             =   1260
         Width           =   1500
      End
      Begin VB.Label LabelMirror 
         BackStyle       =   0  'Transparent
         Caption         =   "Mirror :"
         Height          =   240
         Left            =   30
         TabIndex        =   37
         Top             =   45
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1700
      Left            =   3870
      ScaleHeight     =   1695
      ScaleWidth      =   1905
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1900
      Begin VB.TextBox TextHScale 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   300
         TabIndex        =   41
         Text            =   "100"
         Top             =   330
         Width           =   690
      End
      Begin VB.TextBox TextVScale 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   300
         TabIndex        =   40
         Text            =   "100"
         Top             =   750
         Width           =   690
      End
      Begin VB.VScrollBar VScroll7 
         Height          =   330
         Left            =   990
         Max             =   10000
         Min             =   -10000
         TabIndex        =   39
         Top             =   330
         Value           =   1000
         Width           =   255
      End
      Begin VB.VScrollBar VScroll6 
         Height          =   330
         Left            =   990
         Max             =   10000
         Min             =   -10000
         TabIndex        =   38
         Top             =   765
         Value           =   1000
         Width           =   255
      End
      Begin VB.CommandButton CommandScale 
         Caption         =   "Apply"
         Height          =   315
         Left            =   200
         TabIndex        =   11
         Top             =   1260
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   270
         Index           =   12
         Left            =   1365
         TabIndex        =   45
         Top             =   420
         Width           =   300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   270
         Index           =   11
         Left            =   1365
         TabIndex        =   44
         Top             =   810
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "H :"
         Height          =   225
         Index           =   10
         Left            =   45
         TabIndex        =   43
         Top             =   345
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "V :"
         Height          =   195
         Index           =   9
         Left            =   45
         TabIndex        =   42
         Top             =   780
         Width           =   240
      End
      Begin VB.Label LabelScale 
         BackStyle       =   0  'Transparent
         Caption         =   "Scale :"
         Height          =   240
         Left            =   30
         TabIndex        =   27
         Top             =   45
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1700
      Left            =   2100
      ScaleHeight     =   1695
      ScaleWidth      =   1905
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1900
      Begin VB.VScrollBar VScroll1 
         Height          =   330
         Left            =   1245
         Max             =   1800
         Min             =   -1800
         TabIndex        =   8
         Top             =   495
         Width           =   255
      End
      Begin VB.CommandButton CommandRotate 
         Caption         =   "Apply"
         Height          =   315
         Left            =   200
         TabIndex        =   7
         Top             =   1260
         Width           =   1500
      End
      Begin VB.TextBox TextRotate 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   555
         TabIndex        =   6
         Text            =   "0"
         Top             =   510
         Width           =   690
      End
      Begin VB.Label LabelRotate 
         BackStyle       =   0  'Transparent
         Caption         =   "Rotation :"
         Height          =   240
         Left            =   30
         TabIndex        =   26
         Top             =   45
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "deg"
         Height          =   270
         Index           =   0
         Left            =   1530
         TabIndex        =   18
         Top             =   555
         Width           =   300
      End
      Begin VB.Label Labelangle 
         BackStyle       =   0  'Transparent
         Caption         =   "Angle :"
         Height          =   240
         Left            =   30
         TabIndex        =   9
         Top             =   555
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1700
      Left            =   105
      ScaleHeight     =   1695
      ScaleWidth      =   1905
      TabIndex        =   1
      Top             =   465
      Width           =   1900
      Begin VB.VScrollBar VScroll3 
         Height          =   330
         Left            =   990
         Max             =   2000
         Min             =   -2000
         TabIndex        =   25
         Top             =   765
         Width           =   255
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   330
         Left            =   990
         Max             =   2000
         Min             =   -2000
         TabIndex        =   24
         Top             =   330
         Width           =   255
      End
      Begin VB.TextBox TextVer 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   300
         TabIndex        =   4
         Text            =   "0"
         Top             =   750
         Width           =   690
      End
      Begin VB.TextBox TextHor 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   300
         TabIndex        =   3
         Text            =   "0"
         Top             =   330
         Width           =   690
      End
      Begin VB.CommandButton CommandMove 
         Caption         =   "Apply"
         Height          =   315
         Left            =   200
         TabIndex        =   2
         Top             =   1260
         Width           =   1500
      End
      Begin VB.Label LabelPosition 
         BackStyle       =   0  'Transparent
         Caption         =   "Position :"
         Height          =   240
         Left            =   30
         TabIndex        =   23
         Top             =   45
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "V :"
         Height          =   195
         Index           =   4
         Left            =   45
         TabIndex        =   22
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "H :"
         Height          =   225
         Index           =   3
         Left            =   45
         TabIndex        =   21
         Top             =   345
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "pix"
         Height          =   270
         Index           =   2
         Left            =   1365
         TabIndex        =   20
         Top             =   810
         Width           =   300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "pix"
         Height          =   270
         Index           =   1
         Left            =   1365
         TabIndex        =   19
         Top             =   420
         Width           =   300
      End
   End
End
Attribute VB_Name = "CtrTranform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(c) 2007 - Diomidisk
'Default Property Values:
Const m_def_BackColor = &H8000000F
Const m_def_TypeTransform = 0
'Property Variables:
Dim m_BackColor As OLE_COLOR
Dim m_TypeTransform As Integer

Event TransformScale(X_scale As Single, Y_Scale As Single)
Event TransformSkew(X_Skew As Single, Y_Skew As Single)
Event TransformMove(X_Move As Single, Y_Move As Single)
Event TransformMirror(X_Skew As Integer, Y_Skew As Integer)
Event TransformRotate(t_Angle As Single, xmin As Single, ymin As Single, xmax As Single, ymax As Single)


Private Sub CheckHReflect_Click()
     If CheckHReflect.Value > 0 Or CheckVReflect.Value > 0 Then
        CommandReflect.Enabled = True
     Else
        CommandReflect.Enabled = False
     End If
End Sub


Private Sub CheckVReflect_Click()
      CheckHReflect_Click
End Sub


Private Sub CommandMove_Click()
    Dim X_Move As Single, Y_Move As Single
    
    X_Move = Round(Val(TextHor.Text), 1)
    Y_Move = Round(Val(TextVer.Text), 1)

    RaiseEvent TransformMove(X_Move, Y_Move)
End Sub

Private Sub CommandReflect_Click()
     Dim X_Mirror As Integer, Y_Mirror As Integer
     X_Mirror = CheckVReflect.Value
     Y_Mirror = CheckHReflect.Value
     RaiseEvent TransformMirror(X_Mirror, Y_Mirror)
End Sub

Private Sub CommandRotate_Click()
    
    Dim t_Angle As Single, xmin As Single, ymin As Single, xmax As Single, ymax As Single
    
    t_Angle = Round(Val(TextRotate.Text), 1)
    RaiseEvent TransformRotate(t_Angle, xmin, ymin, xmax, ymax)
End Sub

Private Sub CommandScale_Click()
    Dim ScaleType As Integer, X_scale As Single, Y_Scale As Single
    
    X_scale = Round(Val(TextHScale.Text), 1)
    Y_Scale = Round(Val(TextVScale.Text), 1)
    RaiseEvent TransformScale(X_scale, Y_Scale)
End Sub

Private Sub CommandSkew_Click()
      Dim SkewType As Integer, X_Skew As Single, Y_Skew As Single
      X_Skew = Round(100 + Val(TextHSkew.Text), 1)
      Y_Skew = Round(100 + Val(TextVSkew.Text), 1)
      RaiseEvent TransformSkew(X_Skew, Y_Skew)
End Sub



Private Sub Option1_Click(Index As Integer)
     Select Case Index
     Case 0
        Picture1.Visible = True
        Picture2.Visible = False
        Picture3.Visible = False
        Picture4.Visible = False
        Picture5.Visible = False
        Picture1.SetFocus
     Case 1
        Picture1.Visible = False
        Picture2.Visible = True
        Picture3.Visible = False
        Picture4.Visible = False
        Picture5.Visible = False
        Picture2.SetFocus
     Case 2
        Picture1.Visible = False
        Picture2.Visible = False
        Picture3.Visible = True
        Picture4.Visible = False
        Picture5.Visible = False
        Picture3.SetFocus
     Case 3
        Picture1.Visible = False
        Picture2.Visible = False
        Picture3.Visible = False
        Picture4.Visible = True
        Picture5.Visible = False
        Picture4.SetFocus
     Case 4
        Picture1.Visible = False
        Picture2.Visible = False
        Picture3.Visible = False
        Picture4.Visible = False
        Picture5.Visible = True
        Picture5.SetFocus
     Case Else
        
     End Select
     
End Sub

Private Sub TextHor_Click()
'   If Val(TextHor.Text) > 2000 Or Val(TextHor.Text) < -2000 Then Exit Sub
'      VScroll2.Value = Round(Val(TextHor.Text), 1)
End Sub

Private Sub TextHor_LostFocus()
 If Val(TextHor.Text) > 2000 Or Val(TextHor.Text) < -2000 Then Exit Sub
      VScroll2.Value = Round(Val(TextHor.Text), 1)
End Sub

Private Sub TextHScale_Click()
'      If Val(TextHScale.Text) > 10000 Or Val(TextHScale.Text) < -10000 Then Exit Sub
'      VScroll7.Value = Round(Val(TextHScale.Text) * 10, 1)
End Sub

Private Sub TextHScale_LostFocus()
 If Val(TextHScale.Text) > 10000 Or Val(TextHScale.Text) < -10000 Then Exit Sub
      VScroll7.Value = Round(Val(TextHScale.Text) * 10, 1)
End Sub

Private Sub TextHSkew_Click()
'      If Val(TextHSkew.Text) > 180 Or Val(TextHSkew.Text) < -180 Then Exit Sub
'      VScroll5.Value = Round(Val(TextHSkew.Text) * 10, 1)
End Sub

Private Sub TextHSkew_LostFocus()
If Val(TextHSkew.Text) > 180 Or Val(TextHSkew.Text) < -180 Then Exit Sub
      VScroll5.Value = Round(Val(TextHSkew.Text) * 10, 1)
End Sub

Private Sub TextRotate_LostFocus()
      If Val(TextRotate.Text) > 180 Or Val(TextRotate.Text) < -180 Then Exit Sub
      VScroll1.Value = Round(Val(TextRotate.Text) * 10, 1)
End Sub

Private Sub TextVer_Click()
'      If Val(TextVer.Text) > 2000 Or Val(TextVer.Text) < -2000 Then Exit Sub
'      VScroll3.Value = Round(Val(TextVer.Text), 1)
End Sub

Private Sub TextVer_LostFocus()
 If Val(TextVer.Text) > 2000 Or Val(TextVer.Text) < -2000 Then Exit Sub
      VScroll3.Value = Round(Val(TextVer.Text), 1)
End Sub

Private Sub TextVScale_Click()
'      If Val(TextVScale.Text) > 10000 Or Val(TextVScale.Text) < -10000 Then Exit Sub
'      VScroll6.Value = Round(Val(TextVScale.Text) * 10, 1)
End Sub

Private Sub TextVScale_LostFocus()
If Val(TextVScale.Text) > 10000 Or Val(TextVScale.Text) < -10000 Then Exit Sub
      VScroll6.Value = Round(Val(TextVScale.Text) * 10, 1)
End Sub

Private Sub TextVSkew_Click()
'      If Val(TextVSkew.Text) > 180 Or Val(TextVSkew.Text) < -180 Then Exit Sub
'      VScroll4.Value = Round(Val(TextVSkew.Text) * 10, 1)
End Sub

Private Sub TextVSkew_LostFocus()
    If Val(TextVSkew.Text) > 180 Or Val(TextVSkew.Text) < -180 Then Exit Sub
      VScroll4.Value = Round(Val(TextVSkew.Text) * 10, 1)
End Sub

Private Sub UserControl_Initialize()
      InitializePict
End Sub

Sub InitializePict()
'    Combo1.Clear
'    Combo1.AddItem "Position"
'    Combo1.AddItem "Rotation"
'    Combo1.AddItem "Scale"
'    Combo1.AddItem "Skew"
'    Combo1.AddItem "Mirror"
    
    Picture2.Move Picture1.Left, Picture1.Top, Picture1.Width, Picture1.Height
    Picture3.Move Picture1.Left, Picture1.Top, Picture1.Width, Picture1.Height
    Picture4.Move Picture1.Left, Picture1.Top, Picture1.Width, Picture1.Height
    Picture5.Move Picture1.Left, Picture1.Top, Picture1.Width, Picture1.Height
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_TypeTransform = PropBag.ReadProperty("TypeTransform", m_def_TypeTransform)
    InitializePict
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    Picture1.BackColor = m_BackColor
    Picture2.BackColor = m_BackColor
    Picture3.BackColor = m_BackColor
    Picture4.BackColor = m_BackColor
    Picture5.BackColor = m_BackColor
    UserControl.BackColor = m_BackColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get TypeTransform() As Integer
    TypeTransform = m_TypeTransform
End Property

Public Property Let TypeTransform(ByVal New_TypeTransform As Integer)
    m_TypeTransform = New_TypeTransform
    PropertyChanged "TypeTransform"
    If m_TypeTransform <= Option1.Count Then
       Option1(m_TypeTransform).Value = True
       'Combo1.ListIndex = m_TypeTransform
    End If
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_TypeTransform = m_def_TypeTransform
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("TypeTransform", m_TypeTransform, m_def_TypeTransform)
End Sub

Private Sub VScroll1_Change()
     TextRotate.Text = Format$(VScroll1.Value / 10, "0.0")
End Sub

Private Sub VScroll2_Change()
        TextHor.Text = Format$(VScroll2.Value, "0.0")
End Sub

Private Sub VScroll3_Change()
       TextVer.Text = Format$(VScroll3.Value, "0.0")
End Sub

Private Sub VScroll4_Change()
       TextVSkew.Text = Format$(VScroll4.Value / 10, "0.0")
End Sub

Private Sub VScroll5_Change()
       TextHSkew.Text = Format$(VScroll5.Value / 10, "0.0")
End Sub

Private Sub VScroll6_Change()
     TextVScale.Text = Format$(VScroll6.Value / 10, "0.0")
End Sub

Private Sub VScroll7_Change()
          TextHScale.Text = Format$(VScroll7.Value / 10, "0.0")
End Sub
