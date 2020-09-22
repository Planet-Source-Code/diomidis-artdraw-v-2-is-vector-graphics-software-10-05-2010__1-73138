VERSION 5.00
Begin VB.UserControl CtrColor 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   330
   ScaleWidth      =   2715
End
Attribute VB_Name = "CtrColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(c) 2007 - Diomidisk

'Default Property Values:
Const m_def_ColorFill = 0
Const m_def_ColorBorder = 0
'Property Variables:
Dim m_ColorFill As OLE_COLOR
Dim m_ColorBorder As OLE_COLOR

Private Sub UserControl_Resize()
     Width = 1950
     Height = 525
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = UserControl.Image
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ColorFill() As OLE_COLOR
    ColorFill = m_ColorFill
End Property

Public Property Let ColorFill(ByVal New_ColorFill As OLE_COLOR)
    m_ColorFill = New_ColorFill
    PropertyChanged "ColorFill"
    'Redraw
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ColorBorder() As OLE_COLOR
    ColorBorder = m_ColorBorder
End Property

Public Property Let ColorBorder(ByVal New_ColorBorder As OLE_COLOR)
    m_ColorBorder = New_ColorBorder
    PropertyChanged "ColorBorder"
    'Redraw
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ColorFill = m_def_ColorFill
    m_ColorBorder = m_def_ColorBorder
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_ColorFill = PropBag.ReadProperty("ColorFill", m_def_ColorFill)
    m_ColorBorder = PropBag.ReadProperty("ColorBorder", m_def_ColorBorder)
    Redraw
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ColorFill", m_ColorFill, m_def_ColorFill)
    Call PropBag.WriteProperty("ColorBorder", m_ColorBorder, m_def_ColorBorder)
End Sub

Sub Redraw()
    UserControl.ScaleMode = 3
    Line (5, 12)-(60, 30), m_ColorFill, BF
    Line (5, 12)-(60, 30), 0, B
    Line (64, 12)-(125, 30), m_ColorBorder, BF
    Line (64, 12)-(125, 30), 0, B
    CurrentX = 30: CurrentY = -1:
    Print "Fill"
    CurrentX = 80: CurrentY = -1:
    Print "Border"
    UserControl.ScaleMode = 1
End Sub
