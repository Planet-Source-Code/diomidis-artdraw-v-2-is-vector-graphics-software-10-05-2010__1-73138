VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form Form1 
   Caption         =   "Export file"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   ".."
      Height          =   330
      Left            =   4755
      TabIndex        =   14
      Top             =   2460
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1470
      TabIndex        =   13
      Top             =   2475
      Width           =   3225
   End
   Begin VB.TextBox Text6 
      Height          =   1755
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   3135
      Width           =   5430
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Height          =   555
      Left            =   1560
      TabIndex        =   10
      Top             =   1845
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   945
      TabIndex        =   9
      Text            =   "1"
      Top             =   2475
      Width           =   420
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   270
      Top             =   2295
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Export"
      Height          =   525
      Left            =   1590
      TabIndex        =   8
      Top             =   1215
      Width           =   1755
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   3540
      TabIndex        =   3
      Top             =   375
      Width           =   765
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   2790
      TabIndex        =   2
      Top             =   390
      Width           =   675
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   1695
      TabIndex        =   1
      Top             =   405
      Width           =   660
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   885
      TabIndex        =   0
      Top             =   405
      Width           =   750
   End
   Begin VB.Label Label3 
      Caption         =   "Com"
      Height          =   315
      Left            =   930
      TabIndex        =   11
      Top             =   2145
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y2"
      Height          =   195
      Index           =   1
      Left            =   3645
      TabIndex        =   7
      Top             =   150
      Width           =   195
   End
   Begin VB.Label Label1 
      Caption         =   "X2"
      Height          =   240
      Index           =   1
      Left            =   2820
      TabIndex        =   6
      Top             =   165
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y1"
      Height          =   195
      Index           =   0
      Left            =   1740
      TabIndex        =   5
      Top             =   150
      Width           =   195
   End
   Begin VB.Label Label1 
      Caption         =   "X1"
      Height          =   240
      Index           =   0
      Left            =   915
      TabIndex        =   4
      Top             =   165
      Width           =   585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DataArr() As Variant
Dim ID As Long
Dim SdData() As String

Private Sub Command1_Click()
      Dim X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, ff As Integer
      
      If IsNumeric(Text1.Text) = False Then Beep: Text1.SetFocus: Exit Sub
      If IsNumeric(Text2.Text) = False Then Beep: Text2.SetFocus: Exit Sub
      If IsNumeric(Text3.Text) = False Then Beep: Text3.SetFocus: Exit Sub
      If IsNumeric(Text4.Text) = False Then Beep: Text4.SetFocus: Exit Sub
      X1 = Val(Text1.Text)
      Y1 = Val(Text2.Text)
      X2 = Val(Text3.Text)
      Y2 = Val(Text4.Text)
      
      DataArr = Line1(X1, Y1, X2, Y2)
      
      ff = FreeFile
      Open App.Path + "\Export.txt" For Output As #ff
           Print #ff, "X" + Formatdata(DataArr(0, 0)) + "Y" + Formatdata(DataArr(0, 1)) + "Z" + "0" + "E"
           For i = 0 To UBound(DataArr)
               'Print #ff, "X" + Trim(Str(DataArr(i, 0))) + "Y" + Trim(Str(DataArr(i, 1))) + "Z" + "1" + "E"
               Print #ff, "X" + Formatdata(DataArr(i, 0)) + "Y" + Formatdata(DataArr(i, 1)) + "Z" + "1" + "E"
           Next
           Print #ff, "X" + Formatdata(DataArr(UBound(DataArr), 0)) + "Y" + Formatdata(DataArr(UBound(DataArr), 1)) + "Z" + "0" + "E"
      Close ff
      MsgBox "Export complete", vbInformation
      Text7.Text = App.Path + "\Export.txt"
End Sub

Function Formatdata(dt As Variant) As String
         Dim txt As String
         txt = Format(dt, "00000000.00")
         txt = Replace(txt, ".", "")
         txt = Replace(txt, ",", "")
         Formatdata = txt
End Function

Sub ReadData()
     Dim ff As Integer, txt As String, ids As Long
     Dim fname As String
      ReDim SdData(0)
      ids = 0
      If Text7.Text = "" Then
         Command3_Click
      End If
      fname = Text7.Text
      
      If FileExists(fname) Then
      ff = FreeFile
      Open fname For Input As #ff
           'For i = 1 To UBound(DataArr)
            Do Until EOF(ff)
               Line Input #ff, txt
               ReDim Preserve SdData(ids)
               SdData(ids) = txt
               ids = ids + 1
            Loop
           'Next
     Close ff
     End If
End Sub
Sub SendData()
    If ID > UBound(SdData) Then
       Command2.Enabled = False
       MSComm1.PortOpen = False
       Exit Sub
    End If
    Text6.Text = Text6.Text + SdData(ID) + vbCrLf
    MSComm1.Output = SdData(ID) + vbCr
    ID = ID + 1
End Sub

Function Line1(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Variant

Dim X As Single, Y As Single, dx_Step As Single, dy_Step As Single
Dim dX As Integer, dy As Integer, Steps As Integer, u_ As Integer, a_ As Integer, b_ As Integer
Dim ARR() As Variant
dX = X2 - X1: dy = Y2 - Y1
a_ = Abs(dX): b_ = Abs(dy)
If a_ > b_ Then
   mci# = a_
Else
   mci# = b_
End If

Steps = mci# '+ 1
If Steps = 0 Then Exit Function
dx_Step = dX / Steps: dy_Step = dy / Steps
X = X1 '+ 0.5
Y = Y1 '+ 0.5
ReDim ARR(Steps, 1)

For u_ = 1 To Steps
    'F.PSet (x, y)
    ARR(u_ - 1, 0) = X 'Format(X, "0.000000000")
    ARR(u_ - 1, 1) = Y 'Format(Y, "0.000000000")
    Debug.Print ARR(u_ - 1, 0), ARR(u_ - 1, 1)
    X = X + dx_Step: Y = Y + dy_Step
Next
    'ARR(u_ - 1, 0) = X
    'ARR(u_ - 1, 1) = Y
    ARR(u_ - 1, 0) = X 'Format(X, "0.000000000")
    ARR(u_ - 1, 1) = Y 'Format(Y, "0.000000000")
    
    Line1 = ARR
Debug.Print X, Y
'F.PSet (x, y)

End Function

Private Sub Command2_Click()
          
     MSComm1.CommPort = Val(Text5.Text)
     'MSComm1.Settings = Set1
     MSComm1.InputLen = 0
     MSComm1.Handshaking = comNone 'comRTS
     MSComm1.PortOpen = True
     ReadData
     ID = 0
     SendData
     Command2.Enabled = False
End Sub

Private Sub Command3_Click()
   Dim File_name As String
   Dim sSave As SelectedFile

    FileDialog.sFilter = "Pagra (*.txt)" & Chr$(0) & "*.txt"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_HIDEREADONLY
    FileDialog.sInitDir = App.Path & "\"
    FileDialog.sDefFileExt = "*.txt"
    sSave = ShowOpen(Me.hwnd)
    If Err.Number <> 32755 And sSave.bCanceled = False Then
        Screen.MousePointer = 11
        Text7.Text = sSave.sLastDirectory + sSave.sFile
        File_name = sSave.sFile
'        If DrawControl1.SaveDraw(File_name, File_name) Then
'          ' Update the caption.
'          SetFileName DrawControl1.FileName, DrawControl1.FileTitle
'       End If
       Screen.MousePointer = 0
    End If
        
        
End Sub

Private Sub MSComm1_OnComm()
     Dim Buffer As String, aa As Integer, ff As Long, txtt As String
     On Error GoTo ENDCOM
     
     Select Case MSComm1.CommEvent
     Case comEvReceive
          Buffer = MSComm1.Input
          Text6.Text = Text6.Text + Buffer + vbCrLf
          If InStr(1, Buffer, "*") > 0 Then SendData
          Buffer = ""
          On Error GoTo 0
     End Select
     Exit Sub
ENDCOM:
     Text6.Text = Text6.Text + Error(Err) + vbCrLf
     MsgBox Error(Err), vbCritical
     
End Sub
