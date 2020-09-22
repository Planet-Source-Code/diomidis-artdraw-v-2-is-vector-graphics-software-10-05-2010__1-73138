VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExtReg"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ

Const HKEY_CLASSES_ROOT = &H80000000
Const MAX_PATH = 256&
Const REG_SZ = 1
Const NILL = vbNullString
Const SubKey$ = "shell\open\command"
Const IconKey$ = "DefaultIcon"
Dim RetKey&, lpKey$, IconCmd$, MainKeyHandle As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Const DisplayErrorMsg = False

Public Function Register(ByVal mFileExt As String, _
                         ByVal mFileDescr As String, _
                         ByVal mExeFilePath As String, _
                         Optional ByVal mOverwrite As Boolean = True, _
                         Optional ByVal mIconFile As String = "App") As Boolean
On Error GoTo Fail
Register = False
'error traps / correction
If mExeFilePath = NILL Or mFileExt = NILL Then Exit Function
If Dir(mExeFilePath, 0) = NILL Then Exit Function
If Not Left(mFileExt, 1) = "." Then mFileExt = "." & mFileExt
'overwrite check
If mOverwrite = False Then
    Call RegOpenKey(HKEY_CLASSES_ROOT, mFileExt, RetKey)
    If Not RetKey = 0 Then Exit Function
End If
IconCmd = NILL
'check iconfile
If Dir(mIconFile, 0) = NILL Or mIconFile = NILL Or mIconFile = "App" Then
    mIconFile = mExeFilePath
    IconCmd = ",0"
End If
PS = InStrRev(mExeFilePath, "\")
mExeName = Right(mExeFilePath, Len(mExeFilePath) - PS)
Call RegCreateKey(HKEY_CLASSES_ROOT, mExeName, RetKey)  'create a Root entry for exeFile
Call RegSetValue(RetKey, Empty, REG_SZ, mFileDescr, 0&)
Call RegCreateKey(HKEY_CLASSES_ROOT, mFileExt, RetKey)  'create a Root entry Extension
Call RegSetValue(RetKey, Empty, REG_SZ, mExeName, 0&)
Call RegCreateKey(HKEY_CLASSES_ROOT, mExeName, RetKey)  'set the command line for exeFile
Call RegSetValue(RetKey, SubKey, REG_SZ, mExeFilePath & " %1", MAX_PATH)
Call RegCreateKey(HKEY_CLASSES_ROOT, mExeName, RetKey)  'set default icon for exeFile
Call RegSetValue(RetKey, IconKey, REG_SZ, mIconFile & IconCmd, MAX_PATH)
Register = True
Fail:
Call RegCloseKey(HKEY_CLASSES_ROOT)
Err.Clear
End Function

Public Function UnRegister(ByVal mFileExt As String) As Boolean
On Error GoTo Fail
UnRegister = False
'error traps / correction
If mFileExt = NILL Then Exit Function
If Not Left(mFileExt, 1) = "." Then mFileExt = "." & mFileExt
'read file extension keyvalue
lpKey = String(255, Chr(0))
Call RegQueryValue(HKEY_CLASSES_ROOT, mFileExt, lpKey, RetKey)
mExeName = Left(lpKey, RetKey - 1)
'delete and check if succeeded
Call RegDeleteKey(HKEY_CLASSES_ROOT, mFileExt)
Call RegOpenKey(HKEY_CLASSES_ROOT, mFileExt, RetKey)
If RetKey = 0 Then UnRegister = True
'delete (sub)keys
lpKey = mExeName & "\" & SubKey  'mExeName\shell\open\command
Call RegDeleteKey(HKEY_CLASSES_ROOT, lpKey)
PS = InStrRev(lpKey, "\")        'mExeName\shell\open
lpKey = Left(lpKey, PS - 1)
Call RegDeleteKey(HKEY_CLASSES_ROOT, lpKey)
PS = InStrRev(lpKey, "\")        'mExeName\shell
lpKey = Left(lpKey, PS - 1)
Call RegDeleteKey(HKEY_CLASSES_ROOT, lpKey)
lpKey = mExeName & "\" & IconKey 'mExeName\DefaultIcon
Call RegDeleteKey(HKEY_CLASSES_ROOT, lpKey)
Call RegDeleteKey(HKEY_CLASSES_ROOT, mExeName)
Call RegOpenKey(HKEY_CLASSES_ROOT, mExeName, RetKey)
'check if deteliot succeeded
If UnRegister = True Then UnRegister = IIf(RetKey = 0, True, False)
Fail:
Call RegCloseKey(HKEY_CLASSES_ROOT)
Err.Clear
End Function

Public Function GetBinaryValue(SubKey As String, Entry As String) As String

Dim hKey As Long, rtn As Long
Dim lBufferSize As Long, sBuffer As String
Const REG_BINARY = 3&
Const ERROR_SUCCESS = 0&

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened
      lBufferSize = 1
      rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, 0, lBufferSize) 'get the value from the registry
      sBuffer = Space(lBufferSize)
      rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         GetBinaryValue = sBuffer 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         GetBinaryValue = "Error" 'return Error to the user
         If DisplayErrorMsg = True Then 'if the user wants to errors displayed
            MsgBox ErrorMsg(rtn)  'display the error to the user
         End If
      End If
   Else 'otherwise, if the key couldnt be opened
      GetBinaryValue = "Error" 'return Error to the user
      If DisplayErrorMsg = True Then 'if the user wants to errors displayed
         MsgBox ErrorMsg(rtn)  'display the error to the user
      End If
   End If
End If
 GetBinaryValue = Replace(GetBinaryValue, Chr(0), "")
End Function


Private Sub ParseKey(Keyname As String, Keyhandle As Long)
    
rtn = InStr(Keyname, "\") 'return if "\" is contained in the Keyname

If Left(Keyname, 5) <> "HKEY_" Or Right(Keyname, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
   MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + Keyname 'display error to the user
   Exit Sub 'exit the procedure
ElseIf rtn = 0 Then 'if the Keyname contains no "\"
   Keyhandle = GetMainKeyHandle(Keyname)
   Keyname = "" 'leave Keyname blank
Else 'otherwise, Keyname contains "\"
   Keyhandle = GetMainKeyHandle(Left(Keyname, rtn - 1)) 'seperate the Keyname
   Keyname = Right(Keyname, Len(Keyname) - rtn)
End If

End Sub

Function ErrorMsg(lErrorCode As Long) As String
    
'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

Select Case lErrorCode
       Case 1009, 1015
            GetErrorMsg = "The Registry Database is corrupt!"
       Case 2, 1010
            GetErrorMsg = "Bad Key Name"
       Case 1011
            GetErrorMsg = "Can't Open Key"
       Case 4, 1012
            GetErrorMsg = "Can't Read Key"
       Case 5
            GetErrorMsg = "Access to this key is denied"
       Case 1013
            GetErrorMsg = "Can't Write Key"
       Case 8, 14
            GetErrorMsg = "Out of memory"
       Case 87
            GetErrorMsg = "Invalid Parameter"
       Case 234
            GetErrorMsg = "There is more data than the buffer has been allocated to hold."
       Case Else
            GetErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
End Select

End Function


Function GetMainKeyHandle(MainKeyName As String) As Long

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
   
Select Case MainKeyName
       Case "HKEY_CLASSES_ROOT"
            GetMainKeyHandle = HKEY_CLASSES_ROOT
       Case "HKEY_CURRENT_USER"
            GetMainKeyHandle = HKEY_CURRENT_USER
       Case "HKEY_LOCAL_MACHINE"
            GetMainKeyHandle = HKEY_LOCAL_MACHINE
       Case "HKEY_USERS"
            GetMainKeyHandle = HKEY_USERS
       Case "HKEY_PERFORMANCE_DATA"
            GetMainKeyHandle = HKEY_PERFORMANCE_DATA
       Case "HKEY_CURRENT_CONFIG"
            GetMainKeyHandle = HKEY_CURRENT_CONFIG
       Case "HKEY_DYN_DATA"
            GetMainKeyHandle = HKEY_DYN_DATA
End Select

End Function
