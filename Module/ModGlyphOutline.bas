Attribute VB_Name = "ModGlyphOutline"
' ***************************************************************************************
'     NAME: Modul GlyphOutline
'     DESC: Modul zum Handling der GlyphOutline-API der Windows gdi32.dll
'     DESC: zum ermittlen der Zeichenumrandungen von TrueType Zeichensätzen
' ***************************************************************************************
'
'     AUTHOR:  Stefan Maag
'     CREATE:  12/01/2003
'     CHANGE:
'     COPY:    Der Code geht auf eine Ausschreibung bei Activevb und ein
'     COPY:    Beispiel von Klaus Langbein zurück. Die Idee dazu hatte Robert Christ
'     COPY:    Dieser Code ist frei verfügbar. Die Verwendung des Codes
'     COPY:    oder Codeteilen erfolgt immer auf eigenes Risiko und Verantwortung

' ===========================================================================
'  REM: Quellenangaben:
' ===========================================================================
'  REM:
'  REM:  1: www.activevb.de Beispielcode zu GlyphOutline von Klaus Langbein
'  REM:  2: MSDN Library Juli 2001, Microsoft
'  REM:  3: Das Beispiel in C zu GetGlyphOutline aus der MSDN von Ron Gery
'  REM:     Microsoft Developer Network Technology Group reated: July 10, 1992
'  REM:
' ===========================================================================
'  REM: Weitere benötigte Dateien:
' ===========================================================================
'
'  REM: keine

Option Explicit

Public Enum TT_GlyphFormat
' DESC: Konstanten für den fuFormat-Paramter der GlyphOutline-Funktion
   GGO_BITMAP = 1&     ' DESC: liefert den Zeichenumriss als Bitmap
   GGO_METRICS = 0&    ' DESC: liefert nur die GlyphMetrics Struktur zurück Buffer wird ignoriertEnd Enum
   GGO_NATIVE = 2&     ' DESC: liefert den Zeichenumriss als Array von Eckpunkten, Polygon
End Enum

 
Public Enum TT_CurveType
' DESC: Typ der Punkte im Buffer
   TT_PRIM_LINE = 1&
   TT_PRIM_QSPLINE = 2&
   TT_POLYGON_TYPE = 24&
End Enum

'Public Const GGO_BITMAP = 1      ' DESC: liefert den Zeichenumriss als Bitmap
'Public Const GGO_METRICS = 0     ' DESC: liefert nur die GlyphMetrics Struktur zurück Buffer wird ignoriert
'Public Const GGO_NATIVE = 2      ' DESC: liefert den Zeichenumriss als Array von Eckpunkten, Polygon

'Public Const TT_POLYGON_TYPE = 24
'Public Const TT_PRIM_LINE = 1
'Public Const TT_PRIM_QSPLINE = 2

'  Bei dem Fixed Datentyp handelt es sich um eine 32 Bit Festkommawert
'  mit 16 Bit Ganzzahlanteil und 16 Bit Nachkommanateil. Dies ist
'  jedoch etwas umständlich zu handhaben. Leichter zu verstehen ist
'  es, wenn man diesen Wert als normalen Long betrachtet, der um eine
'  bessere Auflösung der Koordinaten zu bekommen, einfach mit einem
'  Faktor multipliziert ist.
'  Hier ist dieser Faktor 65536, also um 16 Bits nach links geschoben.
'
'  Wir werden diesen Datentyp hier nicht verwenden, sondern die Werte
'  immer als Long betrachten und gegebenfalls durch 65536 teilen und
'  das Ergebnis in einen Gleitpunkt Datentyp (Single/Double) speichern.
'  Dies spart uns viel Ärger!

Type FIXED
    Fract As Integer   ' Nachkommanteil
    Value As Integer   ' Ganzzahlanteil
End Type

Type POINTFX
' DESC: Punkt in der Festkommadefinition der GlyphOutline Funktion
    X As FIXED
    Y As FIXED
End Type

'Type PointAPI
'' DESC: Punkt im Format des Windows API, diesen verwenden wir auch als Ersatz für POINTFX
'    x As Long
'    y As Long
'End Type

Type PointShort
' DESC: Punkt in der Definition mit 16 Bit Integer-Koordinaten
    X As Integer
    Y As Integer
End Type

Type PointSingle
' DESC:  Punkt in der Definition mit (32 Bit) Single-Koordinaten
    X As Single
    Y As Single
End Type

Type GLYPHMETRICS
    gmBlackBoxX As Long    ' Breite der Bitmaps bei fuFormat = CGO_BITMAP
    gmBlackBoxY As Long    ' Höhe der Bitmaps bei fuFormat = CGO_BITMAP
    gmptGlyphOrigin As PointAPI
    gmCellIncX As Integer
    gmCellIncY As Integer
End Type


Type MAT2   ' DESC: die für VB-Programme leichter zu handhabende MAT2 Struktur mit Long
    eM11 As Long
    eM12 As Long
    eM21 As Long
    eM22 As Long
End Type

'Type MAT2          ' DESC: Originaldefintion der MAT2 Struktur mit FIXED Datentyp
'    eM11 As FIXED
'    eM12 As FIXED
'    eM21 As FIXED
'    eM22 As FIXED
'End Type

' DESC: MAT2 ist die Definition für die Transformationsmatrix.
' DESC: Diese Matrix gibt die Grössenverhältnisse und die Ausrichtung,
' DESC: Drehung des erzeugten Zeichenumrisses an.
' DESC:
' DESC: Wir verwenden hier aber nicht die Originaldefinition mit den
' DESC: Werten als FIXED Datentyp sondern wir verwenden wegen besseren
' DESC: Handlings einen einfachen Long. Als Ausgangswert zur Berechung
' DESC: der em-Werte verwenden wir wieder einen Gleitkommawert und
' DESC: multiplizieren diesen mit 65536, speichern diesen in einem Long
' DESC: und erhalten so unseren Festkommawert

' DESC: Die Matrix hat folgendes Aussehen

' DESC: em11 em12
' DESC: em21 em22

Type TTPOLYGONHEADER
    cB As Long             ' Summe der Bytes für Header und Curves
    dwType As Long         ' immer TT_POLYGON_TYPE.
    ' pfxStart As POINTFX   ' erster Punkt als POINTFX Struktur: nicht verwenden
    pfxStart As PointAPI      ' erster Punkt, Punktkoordinaten sind Long
End Type

Type TTPOLYCURVE
    wType As Integer    ' Typ: Polygon oder Bezier
    cpfx As Integer     ' Anzahl der folgenden Punkte incl. dem Startpunkt
    ' apfx As POINTFX    ' Startpunkt der Kurve als POINTFX Struktur: nicht verwenden
    apfx As PointAPI       ' Startpunkt der Kurve, Punktkoordinaten sind Long
End Type

Public Declare Function GetGlyphOutline Lib "gdi32" Alias "GetGlyphOutlineA" _
                  (ByVal hDC As Long, ByVal uChar As Long, _
                   ByVal fuFormat As Long, lpgm As GLYPHMETRICS, _
                   ByVal cbBuffer As Long, lpBuffer As Any, lpmat2 As MAT2) As Long

' Die GlyphOutline Funktion ermittelt die Zeichenumrisse (Outline)

'DWORD GetGlyphOutline(
'  HDC hdc,             // handle to DC
'  UINT uChar,          // character to query
'  UINT uFormat,        // data format
'  LPGLYPHMETRICS lpgm, // glyph metrics
'  DWORD cbBuffer,      // size of data buffer
'  LPVOID lpvBuffer,    // data buffer
'  CONST MAT2 *lpmat2   // transformation matrix
');

Public Const FixedFaktor = 65536    ' DESC: Factor accounting for the FIXED data forma (:= Shift_16)
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal Bytes As Long)
Private Declare Sub MoveMemoryVal Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal Bytes As Long)
Private Declare Sub PeekPoint Lib "msvbvm60.dll" Alias "GetMem8" (Ptr As Any, RetVal As PointAPI)

' ===========================================================================
'  NAME:    GetIdentityMatrix
'  DESC:    Erzeugt die IdentityMatrix, diese ist im Prinzip eine 1 und
'  DESC:    verändert das Zeichen weder in Größe noch Richtung
'  DESC:    man bekommt einen unmaipulierten Glyph
'  RETURN:  IdentityMatrix As MAT2
' ===========================================================================
Public Function GetIdentityMatrix() As MAT2
   With GetIdentityMatrix
      .eM11 = 1 * FixedFaktor
      .eM12 = 0
      .eM21 = 0
      .eM22 = 1 * FixedFaktor
   End With
End Function

' ===========================================================================
'  NAME:    GetShearMatrix
'  DESC:    Erzeugt die horizontale ShearMatrix, Italic Font Simualtion
'  RETURN:  ShearMatrix As MAT2
' ===========================================================================
Public Function GetShearMatrix() As MAT2
   With GetShearMatrix
      .eM11 = 1 * FixedFaktor
      .eM12 = 0
      .eM21 = 0.25 * FixedFaktor
      .eM22 = 1 * FixedFaktor
   End With
End Function

' ===========================================================================
'  NAME:    GetRotationMatrix
'  DESC:    Erzeugt eine Rotationsmatrix basierend auf dem Winkel in Grad
'  RETURN:  Rotationsmatrix As MAT2
' ===========================================================================
Public Function GetRotationMatrix(Angle As Double) As MAT2
   Const PI = 3.14159265358979
   Dim angl As Double
   
   angl = Angle * PI / 180
   With GetRotationMatrix
      .eM11 = (Cos(angl)) * CDbl(FixedFaktor)
      .eM12 = Sin(angl) * CDbl(FixedFaktor)
      .eM21 = -.eM12
      .eM22 = .eM11
   End With
End Function

' ===========================================================================
'  NAME:    GetStrechMatrix
'  DESC:    Erzeugt eine Strechmatrix basierend auf den Steckfaktoren für x/y
'  DESC:    verändert das Zeichen in der Größe
'  RETURN:  StrechMatrix As MAT2
' ===========================================================================
Public Function GetStrechMatrix(ByVal StrechX As Single, ByVal StrechY As Single) As MAT2
   With GetStrechMatrix
      .eM11 = StrechX * FixedFaktor
      .eM12 = 0
      .eM21 = 0
      .eM22 = StrechY * FixedFaktor
   End With
End Function

' ===========================================================================
'  NAME:    GetOutline
'  DESC:    speichert die Punkte der Outline-Kurve in Buffer
'  RETURN:  Anzahl der Bytes in Buffer (nicht der Einträge)
' ===========================================================================
Public Function GetOutline(buffer() As Long, ByVal hDC As Long, ByVal CharASCII As Long, _
                           ByVal fuFormat As TT_GlyphFormat, _
                           metr As GLYPHMETRICS, Matrix As MAT2) As Long
   
   Dim ret As Long
   Dim ByteSize As Long       ' Anzahl der benötigten Bytes im Buffer
   Dim BufSize As Long        ' Einträge im Buffer ( := Ubound(Buffer()) )
   Dim Ptr As Long            ' Pointerwert
   
   ' Beim ersten Aufruf (ByteSize=0) wird die benötigte Länge des Buffers in
   ' Bytes zurückgegen. Oder falls Funktion fehlschlägt der Fehlercode (<0)
   ret = GetGlyphOutline(hDC, CharASCII, fuFormat, metr, ByteSize, ByVal Ptr, Matrix)
   
   If ret > 0 Then
        ByteSize = ret
        BufSize = (ret / 4) - 1  ' /4, da es sich um einen LongBuffer und nicht um Bytes handelt
        ' -1, da von 0 ab dimensioniert wird und BufSize die Anzahl der Einträge darstellt
   Else
        GetOutline = ret       ' Fehlercode zurückgeben
        Exit Function
   End If

   ReDim buffer(BufSize) As Long    ' Buffer in der Benötigten Größe anlegen
   
   Ptr = VarPtr(buffer(0))          ' Startadresse von Buffer()
   
   ' Nun beim 2ten Aufruf wird wirklich der Umriss generiert und in Buffer() gespeichert
   ret = GetGlyphOutline(hDC, CharASCII, fuFormat, metr, ByteSize, ByVal Ptr, Matrix)
   
   GetOutline = ret       ' Returncode zurückgeben
   
   If ret <= 0 Then
      MsgBox "GetGlyphOutline: Error!"
      Exit Function
   End If
   
End Function

' ===========================================================================
'  NAME:    DrawGlyph
'  DESC:    zeichnet den Umriss in die angegeben Picturebox
' ===========================================================================
Public Sub DrawGlyph(buffer() As Long, pb As PictureBox, _
                     ByVal xoff As Long, ByVal yoff As Long, _
                     ByRef Xe() As Long, ByRef Ye() As Long, ByRef TE() As Byte)
  
   Dim i As Long
   Dim j As Long
   Dim Idx As Long
   Dim UB As Long
   Dim EndPoly As Long
   Dim PtsCnt As Long
   Dim ptStart As PointAPI
   Dim X As Single
   Dim Y As Single
   Dim typ As Long
   
   Dim xs() As Long
   Dim ys() As Long
   Dim xp(2) As Long
   Dim yp(2) As Long
   Dim PT() As PointAPI


   UB = UBound(buffer())
'
   ReDim Xe(0)
   ReDim Ye(0)
   ReDim TE(0)
   
   'All polygon header with subordinate PolyCurves through
   Do
      ' --------------------------------------------------------------
      '  Data from the TTPOLGONHEADER - elaborate structure
      ' --------------------------------------------------------------
      
      'Calculation of the last item of the traverse
      'EndPoly = length_Polygon_in_Bytes / 4 + current_buffer_index
      EndPoly = buffer(Idx) \ 4 + Idx
      
      If buffer(Idx + 1) <> TT_POLYGON_TYPE Then
         MsgBox "Polygon errors: curve is not a traverse"
         Exit Sub
      End If
      
      ' Starting point of the polygon train
      ptStart.X = buffer(Idx + 2)
      ptStart.Y = buffer(Idx + 3)
      
      X = ptStart.X / FixedFaktor + xoff
      Y = yoff - ptStart.Y / FixedFaktor
      
      ReDim Preserve Xe(UBound(Xe) + 1)
      ReDim Preserve Ye(UBound(Ye) + 1)
      ReDim Preserve TE(UBound(TE) + 1)
      Xe(UBound(Xe)) = X
      Ye(UBound(Ye)) = Y
      TE(UBound(TE)) = 6
      
      pb.PSet (X, Y), 0
      Idx = Idx + 4
      
      Do
      ' --------------------------------------------------------------
      ' It will begin TTPOLYCURVE structures
      ' --------------------------------------------------------------
         
         ' All curves pass through the polygon
         PtsCnt = buffer(Idx) \ 65536   ' HiWord =Number of the following points
         typ = buffer(Idx) And 65535
         
         Idx = Idx + 1              ' idx on X-coordinate of the first points of the PtsCnt-Punkte
         
         Select Case typ

         Case TT_PRIM_LINE
            ReDim Preserve Xe(UBound(Xe) + PtsCnt)
            ReDim Preserve Ye(UBound(Ye) + PtsCnt)
            ReDim Preserve TE(UBound(TE) + PtsCnt)
            
            For i = 1 To PtsCnt
               X = buffer(Idx) / FixedFaktor + xoff
               Y = yoff - buffer(Idx + 1) / FixedFaktor
               pb.Line -(X, Y)
               
               Xe(UBound(Xe) - PtsCnt + i) = X
               Ye(UBound(Xe) - PtsCnt + i) = Y
               TE(UBound(Xe) - PtsCnt + i) = 2
               Idx = Idx + 2
               ' MsgBox idx
            Next
         
         Case TT_PRIM_QSPLINE
         
            ReDim xs(1 To PtsCnt)
            ReDim ys(1 To PtsCnt)
                
            For i = 1 To PtsCnt
               xs(i) = xoff + buffer(Idx) / FixedFaktor
               ys(i) = yoff - buffer(Idx + 1) / FixedFaktor
               Idx = Idx + 2
            Next i
                
            For i = 1 To PtsCnt - 1
                
               xp(0) = pb.CurrentX
               yp(0) = pb.CurrentY
                                       
               xp(1) = xs(i)
               yp(1) = ys(i)
               
               Select Case PtsCnt - i
               Case 0
                   
               Case 1
                  xp(2) = xs(i + 1)
                  yp(2) = ys(i + 1)
               Case Else
                  xp(2) = xp(1) + (xs(i + 1) - xp(1)) / 2
                  yp(2) = yp(1) + (ys(i + 1) - yp(1)) / 2
               End Select
               
               pb.CurrentX = xp(0)
               pb.CurrentY = yp(0)
                    
               Call QSPLine(5, xp(), yp(), PT())
               
               ReDim Preserve Xe(UBound(Xe) + UBound(PT) + 1)
               ReDim Preserve Ye(UBound(Ye) + UBound(PT) + 1)
               ReDim Preserve TE(UBound(TE) + UBound(PT) + 1)
               
               For j = 0 To UBound(PT)
                   pb.Line -(PT(j).X, PT(j).Y)
                   Xe(UBound(Xe) - UBound(PT) + j) = PT(j).X
                   Ye(UBound(Ye) - UBound(PT) + j) = PT(j).Y
                   TE(UBound(TE) - UBound(PT) + j) = 2
               Next j
               'Erase pt
            Next i
         
         End Select
      ' --------------------------------------------------------------
      ' All these structures TTPOLYCURVE polygon train completed?
      ' --------------------------------------------------------------
      Loop Until Idx >= (EndPoly)
      
       ReDim Preserve Xe(UBound(Xe) + 1)
      ReDim Preserve Ye(UBound(Ye) + 1)
      ReDim Preserve TE(UBound(TE) + 1)
      Xe(UBound(Xe)) = ptStart.X / FixedFaktor + xoff
      Ye(UBound(Ye)) = yoff - ptStart.Y / FixedFaktor
      TE(UBound(TE)) = 3
      
      'now the curve is close to the starting point
      pb.Line -(ptStart.X / FixedFaktor + xoff, yoff - ptStart.Y / FixedFaktor)
      
   ' --------------------------------------------------------------
   'All structures of the TTPOLYGONHEADER Glyph processed?
   ' --------------------------------------------------------------
   Loop Until Idx >= UB 'End of buffer reached?
'   pb.Cls
End Sub


Sub QSPLine(ByVal n As Long, ByRef X() As Long, ByRef Y() As Long, ByRef ptOut() As PointAPI)
        
    Dim i As Long
    Dim t As Double
    Dim tstep As Double
    ReDim ptOut(0 To n)
    
    tstep = (1 / (n))
    
    For i = 0 To n
        t = i * tstep
        ptOut(i).X = (X(0) - 2 * X(1) + X(2)) * t ^ 2 + (2 * X(1) - 2 * X(0)) * t ^ 1 + X(0)
        ptOut(i).Y = (Y(0) - 2 * Y(1) + Y(2)) * t ^ 2 + (2 * Y(1) - 2 * Y(0)) * t ^ 1 + Y(0)
    Next i

End Sub

Public Sub DrawChar(ByVal m_Canvas As PictureBox, ByVal m_TextDraw As String, ByVal m_CurrentX, ByVal m_CurrentY, _
                    ByRef PointCoods() As PointAPI, _
                    ByRef PointType() As Byte, ByRef xmin As Single, ByRef ymin As Single)
                    
      Dim Buf() As Long
      Dim metr As GLYPHMETRICS
      Dim matz As MAT2
      Dim char As Long, FWidth As Single, FHeight As Single, F As Integer
      Dim i As Long, FontWidth As Long, FontHeight As Long
      Dim X1() As Long, Y1() As Long, PT() As Byte
      ReDim PointCoods(0)
      ReDim PointType(0)
      
      FontWidth = m_CurrentX
      FontHeight = m_CurrentY
      
  For F = 1 To Len(m_TextDraw)
      char = Asc(mid(m_TextDraw, F, 1))
      
      CenterText m_Canvas, FontWidth, FontHeight, mid(m_TextDraw, F, 1), _
                 m_Canvas.Font.Size, , , IIf(m_Canvas.FontBold, 800, 400), m_Canvas.Font.Italic, _
                 m_Canvas.Font.Underline, m_Canvas.Font.Strikethrough, m_Canvas.Font.Charset, _
                 , , , , m_Canvas.Font.Name, FWidth, FHeight
      
      FontWidth = FontWidth + FWidth
      'FontHeight = FontHeight + FHeight
      
      matz = GetIdentityMatrix()
      Call GetOutline(Buf(), m_Canvas.hDC, char, GGO_NATIVE, metr, matz)
      DrawGlyph Buf(), m_Canvas, FontWidth, FontHeight, X1(), Y1(), PT() '
      
      ReDim Preserve PointCoods(0 To UBound(PointCoods) + UBound(X1))
      ReDim Preserve PointType(0 To UBound(PointType) + UBound(X1))
      For i = 1 To UBound(X1)
          PointCoods(UBound(PointCoods) - UBound(X1) + i).X = X1(i)
          PointCoods(UBound(PointCoods) - UBound(X1) + i).Y = Y1(i)
          PointType(UBound(PointType) - UBound(PT) + i) = PT(i)
          If xmin > X1(i) Then xmin = X1(i)
          If ymin > Y1(i) Then ymin = Y1(i)
      Next
      
  Next
End Sub


