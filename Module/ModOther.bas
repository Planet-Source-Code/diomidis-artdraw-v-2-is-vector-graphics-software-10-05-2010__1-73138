Attribute VB_Name = "ModOther"
'(c) 2007-2009 DK
Private Declare Function SHPathFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Enum ShellExec
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
End Enum

Public Enum SystemEnv
     s_USERNAME = 1
     s_USERDOMAIN = 2
     s_COMPUTERNAME = 3
     s_LOGONSERVER = 4
     s_TMP = 5
     s_TEMP = 6
     s_SystemRoot = 7
     s_SystemDrive = 8
     s_WinDir = 9
     s_USERPROFILE = 10
     s_APPDATA = 11
     s_CommonProgramFiles = 12
     s_ALLUSERSPROFILE = 13
     s_ProgramFiles = 14
     s_HOMEPATH = 15
     s_HOMEDRIVE = 16
     s_Path = 17
End Enum

Public Function FileExists(Path$) As Boolean
  
   On Error Resume Next
    If Len(Trim(Path$)) = 0 Then FileExists = False: Exit Function
       FileExists = Dir(Trim(Path$), vbNormal) <> ""
       On Error GoTo 0
End Function

Public Function PathExists(sPath As String) As Boolean
    If Len(Environ$("OS")) Then
        PathExists = CBool(SHPathFileExists(StrConv(LTrim$(sPath), vbUnicode)))
    Else
        PathExists = CBool(SHPathFileExists(LTrim$(sPath)))
    End If
End Function

Public Sub SplitPath(FullPath As String, _
                     Optional Drive As String, _
                     Optional Path As String, _
                     Optional Filename As String, _
                     Optional File As String, _
                     Optional Extension As String)
                     
 Dim nPos As Integer
 nPos = InStrRev(FullPath, "\")
 If nPos > 0 Then
   If Left$(FullPath, 2) = "\\" Then
    If nPos = 2 Then
     Drive = FullPath: Path = vbNullString: Filename = vbNullString: File = vbNullString
     Extension = vbNullString
     Exit Sub
    End If
   End If
   Path = Left$(FullPath, nPos - 1)
   Filename = mid$(FullPath, nPos + 1)
   nPos = InStrRev(Filename, ".")
   If nPos > 0 Then
     File = Left$(Filename, nPos - 1)
     Extension = mid$(Filename, nPos + 1)
    Else
     File = Filename
     Extension = vbNullString
   End If
  Else
   nPos = InStrRev(FullPath, ":")
   If nPos > 0 Then
     Path = mid(FullPath, 1, nPos - 1): Filename = mid(FullPath, nPos + 1)
     nPos = InStrRev(Filename, ".")
     If nPos > 0 Then
       File = Left$(Filename, nPos - 1)
       Extension = mid$(Filename, nPos + 1)
      Else
       File = Filename
       Extension = vbNullString
     End If
    Else
     Path = vbNullString: Filename = FullPath
     nPos = InStrRev(Filename, ".")
     If nPos > 0 Then
       File = Left$(Filename, nPos - 1)
       Extension = mid$(Filename, nPos + 1)
      Else
       File = Filename
       Extension = vbNullString
     End If
   End If
 End If
 If Left$(Path, 2) = "\\" Then
   nPos = InStr(3, Path, "\")
   If nPos Then
     Drive = Left$(Path, nPos - 1)
    Else
     Drive = Path
   End If
  Else
   If Len(Path) = 2 Then
    If Right$(Path, 1) = ":" Then
     Path = Path & "\"
    End If
   End If
   If mid$(Path, 2, 2) = ":\" Then
    Drive = Left$(Path, 2)
   End If
 End If
End Sub

'form on top
Public Sub FormOnTop(f_form As Form, i As Boolean)
   
   If i = True Then 'On
      SetWindowPos f_form.hWnd, -1, 0, 0, 0, 0, &H2 + &H1
   Else      'off
      SetWindowPos f_form.hWnd, -2, 0, 0, 0, 0, &H2 + &H1
   End If
   
End Sub

Public Sub SplitRGB(ByVal lColor As Long, ByRef lRed As Long, ByRef lGreen As Long, ByRef lBlue As Long)
   lRed = lColor And &HFF
   lGreen = (lColor And &HFF00&) \ &H100&
   lBlue = (lColor And &HFF0000) \ &H10000
End Sub

'Öoñôþíåé ôï åããñáöü ôïõ áíôéóôïé÷ïõ ðñïãñÜììáôïò ìáæé ìå ôï ðñüãñáììá
Public Function Execute(iFile As String, Optional TypeShow As ShellExec = SW_SHOW) As Boolean
       Dim r As Integer, Msg As String
           afile$ = Left(iFile, InStr(iFile, " "))
       If InStr(1, iFile, "http") = 0 Then
       If Dir$(afile$) = "" Then r = SE_ERR_FNF: GoTo ErrExec
       If InStr(iFile, " ") > 0 Then
          afile$ = Left(iFile, InStr(iFile, " "))
          If FileExists(afile$) = False Then r = SE_ERR_FNF: GoTo ErrExec
       Else
          If FileExists(iFile) = False Then r = SE_ERR_FNF: GoTo ErrExec
       End If
       End If
       r = ShellExecute(0&, vbNullString, iFile, vbNullString, vbNullString, TypeShow)
       
ErrExec:
       If r <= 32 Then
              Select Case r
                  Case SE_ERR_FNF
                      Msg = "File not found"
                  Case SE_ERR_PNF
                      Msg = "Path not found"
                  Case SE_ERR_ACCESSDENIED
                      Msg = "Access denied"
                  Case SE_ERR_OOM
                      Msg = "Out of memory"
                  Case SE_ERR_DLLNOTFOUND
                      Msg = "DLL not found"
                  Case SE_ERR_SHARE
                      Msg = "A sharing violation occurred"
                  Case SE_ERR_ASSOCINCOMPLETE
                      Msg = "Incomplete or invalid file association"
                  Case SE_ERR_DDETIMEOUT
                      Msg = "DDE Time out"
                  Case SE_ERR_DDEFAIL
                      Msg = "DDE transaction failed"
                  Case SE_ERR_DDEBUSY
                      Msg = "DDE busy"
                  Case SE_ERR_NOASSOC
                      Msg = "No association for file extension"
                  Case ERROR_BAD_FORMAT
                      Msg = "Invalid EXE file or error in EXE image"
                  Case Else
                      Msg = "Unknown error"
              End Select
              MsgBox Msg, vbCritical
              Execute = False: Exit Function
          End If
          Execute = True
End Function

'system Environment
Public Function Environment(Env As SystemEnv) As String
        
       Select Case Env
       Case s_USERNAME '= 1
           WindowsTemp = Environ$("USERNAME")
       Case s_USERDOMAIN '= 2
           WindowsTemp = Environ$("USERDOMAIN")
       Case s_COMPUTERNAME ' = 3
           WindowsTemp = Environ$("COMPUTERNAME")
       Case s_LOGONSERVER '= 4
           WindowsTemp = Environ$("LOGONSERVER")
       Case s_TMP '= 5
           WindowsTemp = Environ$("TMP")
       Case s_TEMP '= 6
           WindowsTemp = Environ$("TEMP")
       Case s_SystemRoot ' = 7
           WindowsTemp = Environ$("SystemRoot")
       Case s_SystemDrive ' = 8
           WindowsTemp = Environ$("SystemDrive")
       Case s_WinDir '= 9
           WindowsTemp = Environ$("WinDir")
       Case s_USERPROFILE '= 10
           WindowsTemp = Environ$("USERPROFILE")
       Case s_APPDATA '= 11
           WindowsTemp = Environ$("APPDATA")
       Case s_CommonProgramFiles ' = 12
           WindowsTemp = Environ$("CommonProgramFiles")
       Case s_ALLUSERSPROFILE ' = 13
           WindowsTemp = Environ$("ALLUSERSPROFILE")
       Case s_ProgramFiles ' = 14
           WindowsTemp = Environ$("ProgramFiles")
       Case s_HOMEPATH '= 15
           WindowsTemp = Environ$("HOMEPATH")
       Case s_HOMEDRIVE ' = 16
           WindowsTemp = Environ$("HOMEDRIVE")
       Case s_Path '= 17
           WindowsTemp = Environ$("Path")
       Case Else
           WindowsTemp = ""
       End Select
             
End Function

Public Function IsDebugMode() As Boolean
       If App.LogMode <> 1 Then
          IsDebugMode = True
       End If
End Function


Sub Wait(lDelay As Long)

    'timing loop - measured in milliseconds
    Dim Startl As Long
    Dim SlDelay As Long
    SlDelay = lDelay
    Startl = GetTickCount()
    SlDelay = Startl + SlDelay
    
    Do Until GetTickCount() >= SlDelay Or GetTickCount() < Startl
         DoEvents
    Loop
    
End Sub

'ðáíôá óå ðñùôü ðëÜíï
Public Sub OnForm(f_form As Form, i As Boolean)
   
   If i = True Then 'On
      SetWindowPos f_form.hWnd, -1, 0, 0, 0, 0, &H2 + &H1
   Else      'off
      SetWindowPos f_form.hWnd, -2, 0, 0, 0, 0, &H2 + &H1
   End If
   
End Sub

Public Function PixToMM(frm As Form, mPixels As Single) As Single
       Dim ptr_dpi As Single
       Dim ptr_pixel As Long
       Dim ptr_inch As Single
       
       ptr_pixel = frm.ScaleX(Screen.Width, 1, vbPixels)
       ptr_inch = frm.ScaleX(Screen.Width, 1, vbInches)
       ptr_dpi = ptr_pixel / ptr_inch
       PixToMM = Format((mPixels / ptr_dpi) * 25.4, "0.00")
End Function
