VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Two-dimensional object parent class.

' The object's textual serialization.
Public Serialization As String

' Drawing properties.
Public DrawWidth As Integer
Public DrawStyle As Integer
Public ForeColor As OLE_COLOR
Public FillColor As OLE_COLOR
Public FillColor2 As OLE_COLOR
Public FillMode As FillMode
Public Pattern As String
Public Gradient As Integer
Public FillStyle As Integer
Public TypeDraw As DrawType
Public TextDraw As String
Public AlingText As Integer
Public TypeFill As DrawTypeFill
Public Shade As Boolean
Public ObjLock As Boolean
Public Blend As Integer

'Public TextProperty As New StdFont
Public Angle As Single
Public Bold As Boolean
Public Charset As Integer
Public Italic As Boolean
Public Name As String
Public Size As Integer
Public Strikethrough As Boolean
Public Underline As Boolean
Public Weight As Integer
Public CurrentX As Single
Public CurrentY As Single

Public hRegion As Long
Public Picture As StdPicture

Public EditPoint As Boolean

' Indicates the object is Selected.
Public Selected As Boolean

' The form on which the object is created.
Public canvas As PictureBox

' Clear the object's transformation.
Public Sub ClearTransformation()

End Sub

' Add this transformation to the current one.
Public Sub AddTransformation(m() As Single)

End Sub

' Make transformation .
Public Sub MakeTransformation()
    
End Sub

' Draw the object on the canvas.
Public Sub Draw(ByVal canvas As PictureBox, Optional SetWsc As Boolean = True)

End Sub

' Draw  Select the object on the canvas.
Public Sub DrawSelect(ByVal canvas As PictureBox)

End Sub

' Return this object's bounds.
Public Sub Bound(ByRef xmin As Single, ByRef ymin As Single, ByRef xmax As Single, ByRef ymax As Single)

End Sub

' Return this object's StartPoint.
Public Sub StartBound(ByRef xmin As Single, ByRef ymin As Single)

End Sub

' Draw the object in a metafile.
Public Sub DrawInMetafile(ByVal mf_dc As Long)

End Sub

' Return True if the object is at this location.
Public Function IsAt(ByVal X As Single, ByVal Y As Single) As Boolean

End Function

' Edit new object point.
Public Sub NewPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)

End Sub

' Info from object.
Public Function Info() As String

End Function

' Make point from object
Public Sub ReadPoint(ByRef NumPoints As Long, ByRef X() As Single, ByRef Y() As Single, ByRef TPoint() As Byte)
      
End Sub

' Make point from object
Public Sub ReadTrPoint(ByRef NumPoints As Long, ByRef X() As Single, ByRef Y() As Single, ByRef TPoint() As Byte)
      
End Sub

' Edit new object point.
Public Sub NewTrPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)

End Sub


Public Sub DrawPoint()

End Sub

