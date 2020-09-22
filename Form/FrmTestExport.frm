VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Export Pagra Test"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3960
      Left            =   15
      ScaleHeight     =   260
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   365
      TabIndex        =   0
      Top             =   15
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X() As Single, Y() As Single, Z() As Single
Dim Filename As String
Private Sub Form_Load()
      Me.Show
     
        Filename = Command$
      If FileExists(Filename) Then
         Picture1.ScaleMode = 2
         Picture1_Click
         'DRAW
      Else
        End
      End If
End Sub

Private Sub Form_Resize()
     Picture1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Picture1_Click()
      Dim MSG As String
      If Picture1.ScaleMode = 1 Then
         Picture1.ScaleMode = 2
         MSG = "Tw"
      ElseIf Picture1.ScaleMode = 2 Then
         Picture1.ScaleMode = 3
         MSG = "Pnt"
      ElseIf Picture1.ScaleMode = 3 Then
         Picture1.ScaleMode = 4
         MSG = "Pix"
      ElseIf Picture1.ScaleMode = 4 Then
         Picture1.ScaleMode = 6
         MSG = "In"
      ElseIf Picture1.ScaleMode = 6 Then
         Picture1.ScaleMode = 1
         MSG = "mm"
      End If
      Me.Caption = "Export Pagra " + MSG
      DRAW
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'        Me.Caption = Format(X, "0.00") + " " + Format(Y, "0.00")
End Sub

Private Sub Picture1_Paint()
    DRAW
End Sub


Public Function FileExists(Path$) As Boolean
  
   On Error Resume Next
    If Len(Trim(Path$)) = 0 Then FileExists = False: Exit Function
       FileExists = Dir(Trim(Path$), vbNormal) <> ""
       On Error GoTo 0
End Function
      
Sub DRAW()
ReDim X(0)
     ReDim Y(0)
     ReDim Z(0)
     'MsgBox Command$
     Picture1.Cls
     If FileExists(Filename) Then
     ff = FreeFile
     Open Filename For Input As #ff
          Do Until EOF(ff)
             Line Input #ff, l$ 'Print #ff, "X" + Formatdata(DataArr(0, 0)) + "Y" + Formatdata(DataArr(0, 1)) + "Z" + "0" + "E"
             ReDim Preserve X(UBound(X) + 1)
             ReDim Preserve Y(UBound(Y) + 1)
             ReDim Preserve Z(UBound(Z) + 1)
              X(UBound(X)) = Val(Mid(l$, 2, 10 + 3)) / 100
              Y(UBound(Y)) = Val(Mid(l$, 13, 10 + 3)) / 100
              Z(UBound(Z)) = Val(Mid(l$, 24, 1))
          Loop
           'For i = 0 To UBound(DataArr)
           '    'Print #ff, "X" + Trim(Str(DataArr(i, 0))) + "Y" + Trim(Str(DataArr(i, 1))) + "Z" + "1" + "E"
           '    Print #ff, "X" + Formatdata(DataArr(i, 0)) + "Y" + Formatdata(DataArr(i, 1)) + "Z" + "1" + "E"
           'Next
           'Print #ff, "X" + Formatdata(DataArr(UBound(DataArr), 0)) + "Y" + Formatdata(DataArr(UBound(DataArr), 1)) + "Z" + "0" + "E"
      Close ff
     
      For i = 1 To UBound(X)
          If Z(i) = 0 Then
             Picture1.PSet (X(i), Y(i))
             Picture1.Circle (X(i), Y(i)), 2, vbRed
          Else
             'Picture1.Circle (X(i), Y(i)), 1
             Picture1.Line -(X(i), Y(i))
          End If
      Next
     Else
       End
     End If
End Sub
