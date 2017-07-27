Attribute VB_Name = "Module1"
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Option Explicit
Public x As Long
Public a As Long, b As Long

Public Enum tmsPlatform
    WindowsNT = 1
    Windows = 2
End Enum

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function GetThreadDesktop Lib "user32" (ByVal dwThread As Long) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Public Const SW_HIDE = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOP = 0
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40

Public Const WM_QUERYENDSESSION = &H11
Public Const SMTO_ABORTIFHUNG = &H2
Public Const WM_ENDSESSION = &H16
      
Public Const WM_CLOSE = &H10
Public Const WM_DESTROY = &H2

Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const SE_DEBUG_NAME = "SeDebugPrivilege"

Public Const ANYSIZE_ARRAY = 1
Public Const TOKEN_ADJUST_PRIVILEGES = 32
Public Const TOKEN_QUERY = 8
Public Const SE_PRIVILEGE_ENABLED As Long = 2

Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Type LUID_AND_ATTRIBUTES
    pLuid As LARGE_INTEGER
    Attributes As Long
End Type

Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Public Enum tmsTerminateStatus
    ShutDownWindows = 1
    Restart = 2
    LogOff = 3
End Enum

'Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFileName As String) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long


Public Const VER_PLATFORM_WIN32_NT = 2
Public Const VER_PLATFORM_WIN32_WINDOWS = 1

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long


Public Function StripNulls(OriginalStr As String) As String

If (InStr(OriginalStr, Chr(0)) > 0) Then
   OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
End If

StripNulls = OriginalStr
End Function
 

Public Function AdjProcessPrivilage() As Boolean

Dim lngRet As Long
Dim strMsg As String
Dim strRebootMsg As String
Dim fOkKill
Dim strPlatform As String

Dim ret As Long
Dim hToken As Long
Dim tkp As TOKEN_PRIVILEGES
Dim tkpOld As TOKEN_PRIVILEGES


strPlatform = GetPlatform

If strPlatform = WindowsNT Then
    If OpenProcessToken(GetCurrentProcess(), _
            TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) Then
            
        ret = LookupPrivilegeValue(vbNullString, SE_DEBUG_NAME, tkp.Privileges(0).pLuid)
               
        tkp.PrivilegeCount = 1
        tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
        fOkKill = AdjustTokenPrivileges(hToken, 0, tkp, LenB(tkpOld), tkpOld, ret)
    End If
ElseIf strPlatform = Windows Then
    fOkKill = True
End If
AdjProcessPrivilage = fOkKill

End Function



Public Function EnumWindowProc(ByVal hw As Long, ByVal lParam As Long) As Long
Dim a As Long
Dim strName As String * 255
Dim WinClassBuf As String * 255
Dim lngVis As Long
Dim WinClass As String, WinTitle As String
Dim WinName As String

Dim lRes As Long, lResult As Long
Dim hProcess As Long
Dim ExtCod As Long, ExtCodeProc As Long
Dim ThreadID As Long, ProcessID As Long
Dim TerProc As Long, hHand As Long
Dim RetVal As Long

    'want to display only windows which are visible
    lngVis = IsWindowVisible(hw)
    
    If lngVis > 0 Then
    
        RetVal = GetClassName(hw, WinClassBuf, 255)
        WinClass = StripNulls(WinClassBuf)
        
        If WinClass <> "Shell_TrayWnd" And WinClass <> "Progman" Then
        
            a = GetWindowText(hw, strName, 255)
            WinName = Trim(StripNulls(strName))
            
            If WinName <> "" Then
            
                Form1.Combo1.AddItem WinName
                Form1.cboBkp.AddItem Trim(WinClass)
                Form1.Combo1.ItemData(Form1.Combo1.NewIndex) = hw
                
            End If
            
        End If
        
    End If

EnumWindowProc = True
End Function

Public Sub Destroy()

Dim lRes As Long, lResult As Long
Dim hProcess As Long
Dim ExtCod As Long, ExtCodeProc As Long
Dim ThreadID As Long, ProcessID As Long
Dim TerProc As Long, hHand As Long

ThreadID = GetWindowThreadProcessId(hHand, ProcessID)
hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, ProcessID)
'ExtCod = GetExitCodeProcess(hProcess, ExtCodeProc)
TerProc = TerminateProcess(hProcess, 0) ', ExtCodeProc)

hHand = CloseHandle(hProcess)

End Sub

Public Function GetPlatform() As tmsPlatform

Dim lOSVer As Long
Dim lVer As OSVERSIONINFO
Dim dblSize As Double

lVer.dwOSVersionInfoSize = Len(lVer)
dblSize = GetVersionEx(lVer)
lOSVer = lVer.dwPlatformId

If lOSVer = 1 Then
    GetPlatform = Windows
Else
    GetPlatform = WindowsNT
End If
    
End Function

Public Function g_good(temp1 As String, temp2 As String) As Boolean

Dim first_date As String
Dim second_date As String
Dim month1 As String
Dim day1 As String
Dim year1 As String
Dim month2 As String
Dim day2 As String
Dim year2 As String
Dim found As Boolean
Dim equal As Boolean

found = False
equal = False



first_date = temp1
second_date = temp2

month1 = Month(first_date)
day1 = Month(first_date)
year1 = Month(first_date)

month2 = Month(first_date)
day2 = Month(first_date)
year2 = Month(first_date)

If year1 > year2 Then
    found = True
    ElseIf year1 = year2 Then
        If month1 > month2 Then
            found = True
    ElseIf month1 = month2 Then
        If day1 > day2 Then
            found = True
    ElseIf day1 = day2 Then
        equal = True
        End If
        End If
End If


End Function
