Attribute VB_Name = "mGradFill"
Option Explicit

'©2001/2 Ron van Tilburg  - rivit@f1.net.au
 
'Gradient filling
Private Type TRIVERTEX
  X     As Long
  Y     As Long
  Red   As Integer    'COLOR16
  Green As Integer    'COLOR16
  Blue  As Integer    'COLOR16
  Alpha As Integer    'COLOR16
End Type

Private Type GRADIENT_RECT
  UpperLeft As Long
  LowerRight As Long
End Type

Private Type GRADIENT_TRIANGLE
  Vertex1 As Long
  Vertex2 As Long
  Vertex3 As Long
End Type

Public Const GRADIENT_FILL_RECT_H   As Long = &H0&
Public Const GRADIENT_FILL_RECT_V   As Long = &H1&
Public Const GRADIENT_FILL_TRIANGLE As Long = &H2&

Public Declare Function GradientFill Lib "msimg32" (ByVal hDC As Long, _
                                                    ByVal pVertex As Long, ByVal dwNumVertex As Long, _
                                                    ByVal pMesh As Long, ByVal dwNumMesh As Long, _
                                                    ByVal dwMode As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long

Private Type QTRIVERTEX
  X       As Long
  Y       As Long
  RedLO   As Byte    'COLOR16
  RedHI   As Byte    'COLOR16
  GreenLO As Byte    'COLOR16
  GreenHI As Byte    'COLOR16
  BlueLO  As Byte    'COLOR16
  BlueHI  As Byte    'COLOR16
  AlphaLO As Byte    'COLOR16
  AlphaHI As Byte    'COLOR16
End Type

Private Type QRGB
  R As Byte
  G As Byte
  B As Byte
  z As Byte
End Type

'Calling a routine with 2 Parameters
Public Declare Function VBCallBack2 Lib "user32" Alias "CallWindowProcA" (ByVal FunctionAddress As Long, _
                                                                          ByVal Parameter1 As Long, _
                                                                          ByVal Parameter2 As Long) As Long

Public zBlend As Byte    '0=SOLID,255=Clear      'NOT ACTUALLY USED IT APPEARS
Public Enum GradFills
  GF_RECTHORIZ = 0
  GF_RECTVERT = 1
  GF_TRIFDIAG = 2
  GF_TRIBDIAG = 3
  GF_RECTHORZ2 = 4
  GF_RECTVERT2 = 5
  GF_TRIFDIAG2 = 6
  GF_TRIBDIAG2 = 7
  GF_TRI4WAY = 8
  GF_TRI4WAYB = 9
  GF_TRI4WAYW = 10
  GF_TRI4WAYG = 11
  GF_NGRADFILLS = 12
End Enum

Public Sub GradientFillRectDC(ByVal hDC As Long, _
                              ByVal Left As Long, ByVal Top As Long, _
                              ByVal Right As Long, ByVal Bottom As Long, _
                              ByVal Ink1 As Long, ByVal Ink2 As Long, _
                              ByVal FillType As GradFills)

 Dim Vert(0 To 4)  As QTRIVERTEX
 Dim gRect(0 To 1) As GRADIENT_RECT
 Dim gTri(0 To 3)  As GRADIENT_TRIANGLE
 Dim RGB2 As Long
 
    
  Select Case Abs(FillType) Mod GF_NGRADFILLS
   'RECTANGULAR FILLS
  Case GF_RECTHORIZ, GF_RECTVERT:
        Vert(0).X = Left
        Vert(0).Y = Top
        Call PackVert(Vert(0), Ink1)
    
        Vert(1).X = Right
        Vert(1).Y = Bottom
        Call PackVert(Vert(1), Ink2)
      
        gRect(0).UpperLeft = 0
        gRect(0).LowerRight = 1
    
        Call GradientFill(hDC, VarPtr(Vert(0)), 2, VarPtr(gRect(0)), 1, FillType)
    
   Case GF_RECTHORZ2:
        Vert(0).X = Left
        Vert(0).Y = Top
        Call PackVert(Vert(0), Ink1)
    
        Vert(1).X = (Left + Right) \ 2
        Vert(1).Y = Bottom
        Call PackVert(Vert(1), Ink2)
      
        Vert(2).X = (Left + Right) \ 2
        Vert(2).Y = Top
        Call PackVert(Vert(2), Ink2)
      
        Vert(3).X = Right
        Vert(3).Y = Bottom
        Call PackVert(Vert(3), Ink1)
      
        gRect(0).UpperLeft = 0
        gRect(0).LowerRight = 1
        gRect(1).UpperLeft = 2
        gRect(1).LowerRight = 3
    
        Call GradientFill(hDC, VarPtr(Vert(0)), 4, VarPtr(gRect(0)), 2, GRADIENT_FILL_RECT_H)
    
   Case GF_RECTVERT2:
        Vert(0).X = Left
        Vert(0).Y = Top
        Call PackVert(Vert(0), Ink1)
    
        Vert(1).X = Right
        Vert(1).Y = (Top + Bottom) \ 2
        Call PackVert(Vert(1), Ink2)
      
        Vert(2).X = Left
        Vert(2).Y = (Top + Bottom) \ 2
        Call PackVert(Vert(2), Ink2)
      
        Vert(3).X = Right
        Vert(3).Y = Bottom
        Call PackVert(Vert(3), Ink1)
      
        gRect(0).UpperLeft = 0
        gRect(0).LowerRight = 1
        gRect(1).UpperLeft = 2
        gRect(1).LowerRight = 3
    
        Call GradientFill(hDC, VarPtr(Vert(0)), 4, VarPtr(gRect(0)), 2, GRADIENT_FILL_RECT_V)
    
    'TRIANGULAR FILLS
   Case GF_TRIFDIAG:
'    RGB2 = MixRGB2(Ink1, Ink2)
        Vert(0).X = Left
        Vert(0).Y = Top
        Call PackVert(Vert(0), Ink1)
    
        Vert(1).X = Right
        Vert(1).Y = Top
        Call PackVert(Vert(1), RGB2)
      
        Vert(2).X = Right
        Vert(2).Y = Bottom
        Call PackVert(Vert(2), Ink2)
    
        Vert(3).X = Left
        Vert(3).Y = Bottom
        Call PackVert(Vert(3), RGB2)
          
        gTri(0).Vertex1 = 0
        gTri(0).Vertex2 = 1
        gTri(0).Vertex3 = 3
      
        gTri(1).Vertex1 = 2
        gTri(1).Vertex2 = 3
        gTri(1).Vertex3 = 1
      
        Call GradientFill(hDC, VarPtr(Vert(0)), 4, VarPtr(gTri(0)), 2, GRADIENT_FILL_TRIANGLE)
    
   Case GF_TRIBDIAG:
 '   RGB2 = MixRGB2(Ink1, Ink2)
        Vert(0).X = Left
        Vert(0).Y = Top
        Call PackVert(Vert(0), RGB2)
    
        Vert(1).X = Right
        Vert(1).Y = Top
        Call PackVert(Vert(1), Ink1)
      
        Vert(2).X = Right
        Vert(2).Y = Bottom
        Call PackVert(Vert(2), RGB2)
    
        Vert(3).X = Left
        Vert(3).Y = Bottom
        Call PackVert(Vert(3), Ink2)
          
        gTri(0).Vertex1 = 1
        gTri(0).Vertex2 = 0
        gTri(0).Vertex3 = 2
      
        gTri(1).Vertex1 = 3
        gTri(1).Vertex2 = 2
        gTri(1).Vertex3 = 0
      
        Call GradientFill(hDC, VarPtr(Vert(0)), 4, VarPtr(gTri(0)), 2, GRADIENT_FILL_TRIANGLE)
          
   Case GF_TRIFDIAG2:
        Vert(0).X = Left
        Vert(0).Y = Top
        Call PackVert(Vert(0), Ink1)
    
        Vert(1).X = Right
        Vert(1).Y = Top
        Call PackVert(Vert(1), Ink2)
      
        Vert(2).X = Right
        Vert(2).Y = Bottom
        Call PackVert(Vert(2), Ink1)
    
        Vert(3).X = Left
        Vert(3).Y = Bottom
        Call PackVert(Vert(3), Ink2)
          
        gTri(0).Vertex1 = 0
        gTri(0).Vertex2 = 1
        gTri(0).Vertex3 = 3
      
        gTri(1).Vertex1 = 1
        gTri(1).Vertex2 = 2
        gTri(1).Vertex3 = 3
      
        Call GradientFill(hDC, VarPtr(Vert(0)), 4, VarPtr(gTri(0)), 2, GRADIENT_FILL_TRIANGLE)
   
   Case GF_TRIBDIAG2:
        Vert(0).X = Left
        Vert(0).Y = Top
        Call PackVert(Vert(0), Ink2)
    
        Vert(1).X = Right
        Vert(1).Y = Top
        Call PackVert(Vert(1), Ink1)
      
        Vert(2).X = Right
        Vert(2).Y = Bottom
        Call PackVert(Vert(2), Ink2)
    
        Vert(3).X = Left
        Vert(3).Y = Bottom
        Call PackVert(Vert(3), Ink1)
          
        gTri(0).Vertex1 = 0
        gTri(0).Vertex2 = 1
        gTri(0).Vertex3 = 2
      
        gTri(1).Vertex1 = 2
        gTri(1).Vertex2 = 3
        gTri(1).Vertex3 = 0
      
        Call GradientFill(hDC, VarPtr(Vert(0)), 4, VarPtr(gTri(0)), 2, GRADIENT_FILL_TRIANGLE)
    
   Case GF_TRI4WAY:
        Vert(0).X = Left
        Vert(0).Y = Top
        Call PackVert(Vert(0), Ink2)
    
        Vert(1).X = Right
        Vert(1).Y = Top
        Call PackVert(Vert(1), Ink2)
      
        Vert(2).X = Right
        Vert(2).Y = Bottom
        Call PackVert(Vert(2), Ink2)
    
        Vert(3).X = Left
        Vert(3).Y = Bottom
        Call PackVert(Vert(3), Ink2)
    
        Vert(4).X = (Left + Right) \ 2
        Vert(4).Y = (Top + Bottom) \ 2
        Call PackVert(Vert(4), Ink1)
      
        gTri(0).Vertex1 = 0
        gTri(0).Vertex2 = 1
        gTri(0).Vertex3 = 4
      
        gTri(1).Vertex1 = 1
        gTri(1).Vertex2 = 2
        gTri(1).Vertex3 = 4
      
        gTri(2).Vertex1 = 2
        gTri(2).Vertex2 = 3
        gTri(2).Vertex3 = 4
      
        gTri(3).Vertex1 = 3
        gTri(3).Vertex2 = 0
        gTri(3).Vertex3 = 4
      
        Call GradientFill(hDC, VarPtr(Vert(0)), 5, VarPtr(gTri(0)), 4, GRADIENT_FILL_TRIANGLE)
    
   Case GF_TRI4WAYW:
        Vert(0).X = Left
        Vert(0).Y = Top
        Call PackVert(Vert(0), Ink1)
    
        Vert(1).X = Right
        Vert(1).Y = Top
        Call PackVert(Vert(1), Ink2)
      
        Vert(2).X = Right
        Vert(2).Y = Bottom
        Call PackVert(Vert(2), Ink1)
    
        Vert(3).X = Left
        Vert(3).Y = Bottom
        Call PackVert(Vert(3), Ink2)
    
        Vert(4).X = (Left + Right) \ 2
        Vert(4).Y = (Top + Bottom) \ 2
        Call PackVert(Vert(4), &HFFFFFF)
      
        gTri(0).Vertex1 = 0
        gTri(0).Vertex2 = 1
        gTri(0).Vertex3 = 4
      
        gTri(1).Vertex1 = 1
        gTri(1).Vertex2 = 2
        gTri(1).Vertex3 = 4
      
        gTri(2).Vertex1 = 2
        gTri(2).Vertex2 = 3
        gTri(2).Vertex3 = 4
      
        gTri(3).Vertex1 = 3
        gTri(3).Vertex2 = 0
        gTri(3).Vertex3 = 4
      
        Call GradientFill(hDC, VarPtr(Vert(0)), 5, VarPtr(gTri(0)), 4, GRADIENT_FILL_TRIANGLE)
    
   Case GF_TRI4WAYB:
        Vert(0).X = Left
        Vert(0).Y = Top
        Call PackVert(Vert(0), Ink1)
    
        Vert(1).X = Right
        Vert(1).Y = Top
        Call PackVert(Vert(1), Ink2)
      
        Vert(2).X = Right
        Vert(2).Y = Bottom
        Call PackVert(Vert(2), Ink1)
    
        Vert(3).X = Left
        Vert(3).Y = Bottom
        Call PackVert(Vert(3), Ink2)
    
        Vert(4).X = (Left + Right) \ 2
        Vert(4).Y = (Top + Bottom) \ 2
        Call PackVert(Vert(4), &H0&)
      
        gTri(0).Vertex1 = 0
        gTri(0).Vertex2 = 1
        gTri(0).Vertex3 = 4
      
        gTri(1).Vertex1 = 1
        gTri(1).Vertex2 = 2
        gTri(1).Vertex3 = 4
      
        gTri(2).Vertex1 = 2
        gTri(2).Vertex2 = 3
        gTri(2).Vertex3 = 4
          
        gTri(3).Vertex1 = 3
        gTri(3).Vertex2 = 0
        gTri(3).Vertex3 = 4
          
        Call GradientFill(hDC, VarPtr(Vert(0)), 5, VarPtr(gTri(0)), 4, GRADIENT_FILL_TRIANGLE)
    
   Case GF_TRI4WAYG:
        Vert(0).X = Left
        Vert(0).Y = Top
        Call PackVert(Vert(0), &H0)
    
        Vert(1).X = Right
        Vert(1).Y = Top
        Call PackVert(Vert(1), Ink2)
    
        Vert(2).X = Right
        Vert(2).Y = Bottom
        Call PackVert(Vert(2), &HFFFFFF)
    
        Vert(3).X = Left
        Vert(3).Y = Bottom
        Call PackVert(Vert(3), Ink1)
        
        Vert(4).X = (Left + Right) \ 2
        Vert(4).Y = (Top + Bottom) \ 2
        Call PackVert(Vert(4), ((Ink1 And &HFF0000) + (Ink2 And &HFF0000)) \ 2 _
                            Or ((Ink1 And &HFF00&) + (Ink2 And &HFF00&)) \ 2 _
                            Or ((Ink1 And &HFF&) + (Ink2 And &HFF&)) \ 2)
          
        gTri(0).Vertex1 = 0
        gTri(0).Vertex2 = 1
        gTri(0).Vertex3 = 4
          
        gTri(1).Vertex1 = 1
        gTri(1).Vertex2 = 2
        gTri(1).Vertex3 = 4
          
        gTri(2).Vertex1 = 2
        gTri(2).Vertex2 = 3
        gTri(2).Vertex3 = 4
          
        gTri(3).Vertex1 = 3
        gTri(3).Vertex2 = 0
        gTri(3).Vertex3 = 4
          
        Call GradientFill(hDC, VarPtr(Vert(0)), 5, VarPtr(gTri(0)), 4, GRADIENT_FILL_TRIANGLE)
  End Select
  
End Sub

Private Sub PackVert(ByRef Vert As QTRIVERTEX, ByRef Ink As Long)

  Dim Color As QRGB
  
  Call CopyMemory(Color, Ink, 4)
  Vert.RedHI = Color.R
  Vert.GreenHI = Color.G
  Vert.BlueHI = Color.B
  Vert.AlphaHI = zBlend

End Sub
    
':) Ulli's VB Code Formatter V2.6.10 (20-Dec-01 15:47:01) 29 + 275 = 304 Lines
