Attribute VB_Name = "Export_Txt"
Public ExportTxt As Boolean
Type dexp
     Xe() As Single
     Ye() As Single
     Ze() As Integer
End Type

Public Obj() As dexp

Sub ExportOn(OnOff As Boolean)
    ExportTxt = OnOff
    ReDim Obj(0)
    ReDim Obj(0).Xe(0)
    ReDim Obj(0).Ye(0)
    ReDim Obj(0).Ze(0)
End Sub

Sub AddExportPoint(pic As Object, PT() As PointAPI, TP() As Byte)
    
    Dim X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single, X4 As Single, Y4 As Single
    Dim FP As Long, LP As Long, FP1 As Long, i As Long, XT As Single, YT As Single, FindOneMore As Boolean
    Dim Arr() As Variant, id As Long
    
'   On Error Resume Next
    
    FindOneMore = False
    
    FP1 = UBound(Obj(id).Xe)
    ReDim Preserve Obj(UBound(Obj) + 1)
    id = UBound(Obj)
    ReDim Preserve Obj(id).Xe(0)
    ReDim Preserve Obj(id).Ye(0)
    ReDim Preserve Obj(id).Ze(0)
  
        
    i = 1
    Do
        If i + 1 >= UBound(PT) And FP1 > 0 Then
           Exit Do
        End If
        
        X1 = PT(i).X
        Y1 = PT(i).Y
        
        If i + 1 > UBound(PT) Then
           Exit Do
        End If
        X2 = PT(i + 1).X
        Y2 = PT(i + 1).Y
        
        If TP(i + 1) <> 4 And TP(i + 1) <> 6 Then
           Arr = Line1(X1, Y1, X2, Y2)
           i = i + 1
        '   Debug.Print X1, X2, Y1, Y2
        ElseIf TP(i) = 3 And i > 1 And i < UBound(TP) Then
           ReDim Arr(1, 1)
           Arr(0, 0) = X1
           Arr(0, 1) = Y1
           Arr(1, 0) = X2
           Arr(1, 1) = Y2
           XT = X2
           YT = Y2
           FindOneMore = True
           i = i + 1
           'Stop
        Else
          If i + 2 > UBound(PT) Then
             Exit Do
          End If
           X3 = PT(i + 2).X
           Y3 = PT(i + 2).Y
           X4 = PT(i + 3).X
           Y4 = PT(i + 3).Y
           Arr = DrawBezier1(X1, Y1, X2, Y2, X3, Y3, X4, Y4)
           i = i + 3
           Debug.Print X1, X2, X3, X4, Y1, Y2, Y3, Y4
        End If
        If UBound(Arr) > 0 Then
        FP = UBound(Obj(id).Xe)
        LP = UBound(Arr)
        
        ReDim Preserve Obj(id).Xe(UBound(Obj(id).Xe) + LP + 1)
        ReDim Preserve Obj(id).Ye(UBound(Obj(id).Ye) + LP + 1)
        ReDim Preserve Obj(id).Ze(UBound(Obj(id).Ze) + LP + 1)
        
        For t = 0 To LP
           Obj(id).Xe(FP + t + 1) = Arr(t, 0)
           Obj(id).Ye(FP + t + 1) = Arr(t, 1)
           If i - 1 > 0 Then
             If (TP(i - 1) <> 4 And TP(i) = 6) And (TP(i) = 6 And i > 1 And i < UBound(TP)) Then
                Obj(id).Ze(FP + t + 1) = 0
             Else
                Obj(id).Ze(FP + t + 1) = 1
             End If
           Else
              Obj(id).Ze(FP + t + 1) = 1
           End If
        Next
        End If
     Loop
     
AD1:
    If FindOneMore Then
       Arr = Line1(XT, YT, X2, Y2)
       FP = UBound(Obj(id).Xe)
       LP = UBound(Arr)
       If LP > 0 Then
        ReDim Preserve Obj(id).Xe(UBound(Obj(id).Xe) + LP)
        ReDim Preserve Obj(id).Ye(UBound(Obj(id).Ye) + LP)
        ReDim Preserve Obj(id).Ze(UBound(Obj(id).Ze) + LP)
        For t = 0 To LP
           Obj(id).Xe(FP + t) = Arr(t, 0)
           Obj(id).Ye(FP + t) = Arr(t, 1)
           If i - 2 > 0 Then
             If (TP(i - 1) <> 4 And TP(i) = 6) And (TP(i) = 6 And i > 1 And i < UBound(TP)) Then
                Obj(id).Ze(FP + t) = 0
                FindOneMore = False
             Else
                Obj(id).Ze(FP + t) = 1
             End If
           Else
              Obj(id).Ze(FP + t) = 1
           End If
        Next
       End If
    End If
    
    Obj(id).Xe(0) = Obj(id).Xe(1)
    Obj(id).Ye(0) = Obj(id).Ye(1)
    Obj(id).Ze(0) = 0
    
    ReDim Preserve Obj(id).Xe(UBound(Obj(id).Xe) + 1)
    ReDim Preserve Obj(id).Ye(UBound(Obj(id).Ye) + 1)
    ReDim Preserve Obj(id).Ze(UBound(Obj(id).Ze) + 1)
    Obj(id).Xe(UBound(Obj(id).Xe)) = Obj(id).Xe(UBound(Obj(id).Xe) - 1)
    Obj(id).Ye(UBound(Obj(id).Ye)) = Obj(id).Ye(UBound(Obj(id).Ye) - 1)
    Obj(id).Ze(UBound(Obj(id).Ze)) = 0
    
    On Error GoTo 0
End Sub

Function SaveExportTxt(fform As Form, Filename As String)
      Dim ff As Integer, id As Integer, i As Integer
      ff = FreeFile
      Open Filename For Output As #ff
           For id = 1 To UBound(Obj)
             For i = 0 To UBound(Obj(id).Xe)
               X = fform.ScaleX(Obj(id).Xe(i), vbPixels, gScaleMode)
               Y = fform.ScaleX(Obj(id).Ye(i), vbPixels, gScaleMode)
               Print #ff, "X" + FormatData(X) + "Y" + FormatData(Y) + "Z" + Trim(Str(Obj(id).Ze(i))) + "E"
             Next
           Next
      Close ff
      
End Function

Private Function FormatData(dt As Variant) As String
       Dim txt As String
       txt = Format(dt, "00000000.00")
       txt = Replace(txt, ".", "")
       txt = Replace(txt, ",", "")
       FormatData = txt
End Function

Private Function Line1(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Variant

    Dim X As Single, Y As Single, dx_Step As Single, dy_Step As Single
    Dim dX As Integer, dy As Integer, Steps As Integer, u_ As Integer, a_ As Integer, b_ As Integer
    Dim Arr() As Variant
    
    dX = X2 - X1: dy = Y2 - Y1
    a_ = Abs(dX): b_ = Abs(dy)
    If a_ > b_ Then
        mci# = a_
    Else
        mci# = b_
    End If
    
    Steps = mci# '+ 1
    
    If (dX = 0 And Steps = 0) Or (dy = 0 And Steps = 0) Then
       ReDim Arr(0, 0)
       GoTo L1: 'Exit Function
    End If
    
    dx_Step = dX / Steps: dy_Step = dy / Steps
    X = X1 '+ 0.5
    Y = Y1 '+ 0.5
    ReDim Arr(Steps, 1)

    For u_ = 1 To Steps
        'F.PSet (x, y)
        Arr(u_ - 1, 0) = X 'Format(X, "0.000000000")
        Arr(u_ - 1, 1) = Y 'Format(Y, "0.000000000")
        'Debug.Print Arr(u_ - 1, 0), Arr(u_ - 1, 1)
        X = X + dx_Step: Y = Y + dy_Step
    Next
    'ARR(u_ - 1, 0) = X
    'ARR(u_ - 1, 1) = Y
    Arr(u_ - 1, 0) = X 'Format(X, "0.000000000")
    Arr(u_ - 1, 1) = Y 'Format(Y, "0.000000000")
L1:
    Line1 = Arr
    'Debug.Print X, Y
    
End Function

Function DrawBezier1(X1 As Single, Y1 As Single, _
                     X2 As Single, Y2 As Single, _
                     X3 As Single, Y3 As Single, _
                     X4 As Single, Y4 As Single, _
                     Optional du As Single = 0.005) As Variant
    ' Draws a Bezier curve using the control points given in  Cont(...). Uses delta u steps
    Dim Cnt(3, 2), Arr() As Variant, aa As Long
    Dim bv As Variant
    
    'n = nc - 1 'N = number of control points -1
    n = 3
    Cnt(0, 0) = X1
    Cnt(0, 1) = Y1
    Cnt(1, 0) = X2
    Cnt(1, 1) = Y2
    Cnt(2, 0) = X3
    Cnt(2, 1) = Y3
    Cnt(3, 0) = X4
    Cnt(3, 1) = Y4
    
    'picDisplay.PSet (Cnt(0, 0), Cnt(0, 1)) 'Plot the first point
    ReDim Arr(1 / du + 1, 1)
    aa = 0
    For u = 0 To 1 Step du
        X = 0: Y = 0
        For k = 0 To n ' For Each control point
            bv = b(k, n, u) ' Calculate blending Function
            X = X + Cnt(k, 0) * bv
            Y = Y + Cnt(k, 1) * bv
        Next k
        Arr(aa, 0) = X
        Arr(aa, 1) = Y
        aa = aa + 1
        'picDisplay.Line -(X, Y), 0  ' Draw To the point
        'picDisplay.PSet (X, Y)
    Next u
    Arr(UBound(Arr), 0) = Cnt(3, 0)
    Arr(UBound(Arr), 1) = Cnt(3, 1)
    
    For u = 0 To UBound(Arr) - 1
        If Abs(Arr(u, 0) - Arr(u + 1, 0)) > 1 Then
          ' Stop
        End If
    Next
    DrawBezier1 = Arr
    'picDisplay.Line -(Cont(n, 0), Cont(n, 1)), 65535
    'picDisplay.PSet (Cnt(n, 0), Cnt(n, 1))
End Function

Private Function b(k, n, u)
    'Bezier blending function
    b = C(n, k) * (u ^ k) * (1 - u) ^ (n - k)
    
End Function

Private Function C(n, r)
    ' Implements c!/r!*(n-r)!
    C = fact(n) / (fact(r) * fact(n - r))
    
End Function

Private Function fact(n)
    ' Recursive factorial fucntion
    If n = 1 Or n = 0 Then
        fact = 1
    Else
        fact = n * fact(n - 1)
    End If
End Function

