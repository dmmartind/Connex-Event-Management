VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connex Event Management System"
   ClientHeight    =   4980
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7965
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboBkp 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmMain.frx":0442
      Left            =   360
      List            =   "frmMain.frx":0444
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   1350
      Left            =   5760
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   1815
      Left            =   1920
      OleObjectBlob   =   "frmMain.frx":0446
      TabIndex        =   0
      Top             =   3120
      Width           =   3975
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6480
      OleObjectBlob   =   "frmMain.frx":0630
      Top             =   240
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu utilities 
      Caption         =   "&Admin"
      Begin VB.Menu employee 
         Caption         =   "Employee View/Entry"
      End
      Begin VB.Menu client 
         Caption         =   "Client View/Entry"
      End
      Begin VB.Menu room 
         Caption         =   "Room View/Entry"
      End
      Begin VB.Menu activities 
         Caption         =   "Activities View/Entry"
      End
      Begin VB.Menu Schedule 
         Caption         =   "Schedule View/Entry"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu user_group 
         Caption         =   "User and Group Accounts"
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu make_backup 
         Caption         =   "Backup TSR Setup"
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu settings 
         Caption         =   "Settings"
      End
   End
   Begin VB.Menu reports_m 
      Caption         =   "Reports"
      Begin VB.Menu event_m 
         Caption         =   "Event"
      End
      Begin VB.Menu schedule_m 
         Caption         =   "Schedule"
      End
   End
   Begin VB.Menu employee_action 
      Caption         =   "&Employee Action"
      Begin VB.Menu clockin 
         Caption         =   "Clock In"
      End
      Begin VB.Menu clockout 
         Caption         =   "Clock Out"
      End
      Begin VB.Menu space3 
         Caption         =   "-"
      End
      Begin VB.Menu view_schedule 
         Caption         =   "View Schedule"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu contents 
         Caption         =   "Contents"
      End
      Begin VB.Menu unlock_sp 
         Caption         =   "Unlock"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Option Explicit


Public LoginSucceeded As Boolean
Public directory1 As String
Public connection_string As String
Public connection_string2 As String
Public user_name1 As String
Public setting_temp As Integer
Public voice As Integer
Dim hConsole As Long
Dim status As Boolean



Private Sub about_Click()
On Error GoTo has_error
'about1.Show vbModal
about1.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub activities_Click()
On Error GoTo has_error
Dialog1.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub client_Click()
On Error GoTo has_error
Form9.Show vbModal
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub clockin_Click()

On Error GoTo has_error
check_in.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"

Unload Me
End Sub

Private Sub clockout_Click()




On Error GoTo has_error
Form3.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub


Private Sub create_Click()

Rem Form10.Show vbModal
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub employee_Click()
On Error GoTo has_error
Form14.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub event_m_Click()
On Error GoTo has_error
Form5.Show vbModal
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub exit_Click()
On Error GoTo has_error
Dim OneForm As Form
Dim count As Integer
NeMos_deac

'////////////////////////////////////////////////////////////////////////////////////////////////////////////
For Each OneForm In Forms
Unload OneForm
Next
'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Exit Sub
'errhand:
'MsgBox Err.Description, vbOKOnly + vbCritical

'//////////////////////////////////////////////////////////////////////////////////////
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub Form_Load()

On Error GoTo has_error
Skin1.ApplySkin Me.hwnd


Dim temp As Long
Dim temp2 As String
Dim temp3 As Long
Dim strfilepath As String
Dim temp5 As String
Dim lngKeyHandle As String



event_m.Enabled = False
employee.Enabled = False
client.Enabled = False
room.Enabled = False
Schedule.Enabled = False
user_group.Enabled = False
make_backup.Enabled = False
activities.Enabled = False
clockin.Enabled = False
clockout.Enabled = False
view_schedule.Enabled = False
about.Enabled = False
Contents.Enabled = False
settings.Enabled = True
schedule_m.Enabled = False

temp5 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\My Documents\temp.mde;Jet OLEDB:Database Password=51289;"



unlock_con
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"

Unload Me

End Sub

Private Sub make_backup_Click()
On Error GoTo has_error
Form7.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub room_Click()
On Error GoTo has_error
Form6.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub
Private Sub schedule_Click()
On Error GoTo has_error
Form2.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub



Private Sub schedule_m_Click()
On Error GoTo has_error
room_report.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub settings_Click()

On Error GoTo has_error

Dim strfilepath As String
Settings1.Show vbModal
If setting_temp = 1 Then
connection_string = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & directory1 & "\temp.mde;Jet OLEDB:Database Password=V![4J5hS~-2TmbXn18@2;"
connection_string2 = "Provider=SQLOLEDB.1;Data Source=" & directory1 & "\temp.mde;Jet OLEDB:Database Password=51289;"
strfilepath = App.Path & "\approxy.dat"
Open strfilepath For Output As #1
Write #1, connection_string
Write #1, connection_string2
Close #1
End If
'///////////////////////////////////////////////////////////////////////////////////////////////
On Error GoTo h_run

Dim adn1 As ADODB.Connection
Dim data1 As ADODB.Recordset
Dim connectionstring As String

Set adn1 = New ADODB.Connection
Set data1 = New ADODB.Recordset

With adn1
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data1
.ActiveConnection = adn1
.CursorType = adOpenDynamic
.LockType = adLockOptimistic
.Source = "SELECT * FROM grub"
End With


With data1
.Open
While .EOF = False
.Delete
.MoveNext
Wend

.AddNew
.Fields("Current Directory") = directory1
.Update
.Close
End With

adn1.Close

Exit Sub


h_run:
MsgBox "Database was not found", vbExclamation, "Connex Event Management"

'///////////////////////////////////////////////////////////////////////////////////////////////
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub unlock_sp_Click()
Rem frmLogin1.Show vbModal
Rem If LoginSucceeded = True Then
Rem go_com
Rem Else
Rem End
Rem End If

End Sub

Private Sub user_group_Click()
On Error GoTo has_error
Form4.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub view_schedule_Click()
On Error GoTo has_error
Form20.Show vbModal
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub
Private Sub go_com()
On Error GoTo has_error
MsgBox "This License expires in 30 days", vbOKOnly, "Connex Event Management"
employee.Enabled = True
client.Enabled = True
room.Enabled = True
Schedule.Enabled = True
user_group.Enabled = False
'Create.Enabled = False
make_backup.Enabled = True

clockin.Enabled = True
clockout.Enabled = False
view_schedule.Enabled = False
about.Enabled = True
Contents.Enabled = False
settings.Enabled = False
unlock_sp.Enabled = False
Rem Dim c As New cRegistry

    Rem         With c
    Rem        .ClassKey = HKEY_LOCAL_MACHINE
    Rem        .SectionKey = "david\0001\Enum\Bios\TGAN-1586CEZZ\PG-C30XU"
    Rem        .ValueKey = "95434"
    Rem        .ValueType = REG_SZ
    Rem         .Value = "78-6970-9090-6"
    Rem        End With
         
 
  Rem       .ClassKey = HKEY_LOCAL_MACHINE
  Rem       .SectionKey = "Config\0001\Enum\Bios\TGAN-1586CEZZ\PG-C30XU"
  Rem       .ValueKey = "IDF41"
  Rem       .ValueType = REG_DWORD
  Rem       .Value = 305264
  Rem       End With

unlock_con

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

         
End Sub



Private Sub unlock_con()

On Error GoTo has_error


status = True
Dim strfilepath As String
Dim temp5 As String

employee.Enabled = False
client.Enabled = False
room.Enabled = False
Schedule.Enabled = False
user_group.Enabled = False
make_backup.Enabled = False
clockin.Enabled = False
clockout.Enabled = False
view_schedule.Enabled = False
about.Enabled = False
Contents.Enabled = False
activities.Enabled = False
event_m.Enabled = False
schedule_m.Enabled = False

settings.Enabled = True

unlock_sp.Enabled = False

strfilepath = App.Path & "\approxy.dat"

On Error GoTo HandleErrors

Open strfilepath For Input As #1
Input #1, temp5
If temp5 <> "" Then
connection_string = temp5
temp5 = ""
Input #1, temp5
If temp5 <> "" Then
connection_string2 = temp5
Close #1
frmLogin.Show vbModal
Else
status = False
MsgBox "This program could not connect to a database. Please check your setting to make sure an address has been set.", vbCritical, "Connex Event Management"
Close #1
Open strfilepath For Output As #2
Write #2, ""
Close #2
End If
Else
status = False
MsgBox "This program could not connect to a database. Please check your setting to make sure a directory has been set.", vbCritical, "Connex Event Management"
Close #1
End If

'///////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////
Dim adn1 As ADODB.Connection
Dim data1 As ADODB.Recordset

Set adn1 = New ADODB.Connection
Set data1 = New ADODB.Recordset



With adn1
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data1
.ActiveConnection = adn1
.CursorType = adOpenDynamic
.LockType = adLockOptimistic
.Source = "SELECT * FROM backup"
End With

Dim sTemp As String


With data1
.Open
If .EOF = False Then
If ![Power Switch] = 1 Then
sTemp = App.Path & "\btsr.exe"

If sTemp = "" Then Exit Sub

Shell sTemp, vbNormalFocus

End If
End If
.Close
End With
adn1.Close

'/////////////////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////////////////////////////////
Exit Sub

Form_Load_Exit:
Exit Sub

HandleErrors:
Close #1
MsgBox "Error:: Database could not be detected", vbCritical, "Connex Event Management"
MsgBox "This program could not connect to a database. Please check your setting to make sure an address has been set.", vbCritical, "Connex Event Management"
Open strfilepath For Output As #2
Write #2, ""
Close #2

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub
Public Sub unlock_admin2()

On Error GoTo has_error
Dim strfilepath As String
Dim temp5 As String

event_m.Enabled = True
schedule_m.Enabled = True
activities.Enabled = True
employee.Enabled = True
client.Enabled = True
room.Enabled = True
Schedule.Enabled = True
user_group.Enabled = True
make_backup.Enabled = True
clockin.Enabled = True
clockout.Enabled = True
view_schedule.Enabled = True
about.Enabled = True
Contents.Enabled = False
settings.Enabled = True

unlock_sp.Enabled = False

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"

Unload Me

End Sub



Public Sub unlock_admin()

On Error GoTo has_error
Dim strfilepath As String
Dim temp5 As String

event_m.Enabled = True
schedule_m.Enabled = True
activities.Enabled = True
employee.Enabled = True
client.Enabled = True
room.Enabled = True
Schedule.Enabled = True
user_group.Enabled = False
make_backup.Enabled = False
clockin.Enabled = True
clockout.Enabled = True
view_schedule.Enabled = True
about.Enabled = True
Contents.Enabled = False
settings.Enabled = True

unlock_sp.Enabled = False

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub



Public Sub unlock_user()

On Error GoTo has_error
Dim strfilepath As String
Dim temp5 As String

employee.Enabled = False
client.Enabled = False
room.Enabled = False
Schedule.Enabled = False
user_group.Enabled = False

make_backup.Enabled = False
clockin.Enabled = True
clockout.Enabled = True
view_schedule.Enabled = True
about.Enabled = True
Contents.Enabled = False
settings.Enabled = True

unlock_sp.Enabled = False

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"

Unload Me

End Sub

Private Sub Destroy(ByVal Hwd As Long)
On Error GoTo has_error

Dim lRes As Long, lResult As Long

Dim hProcess As Long
Dim ExtCod As Long, ExtCodeProc As Long
Dim ThreadID As Long, ProcessID As Long
Dim TerProc As Long, hHand As Long

ThreadID = GetWindowThreadProcessId(Hwd, ProcessID)
hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, ProcessID)
'ExtCod = GetExitCodeProcess(hProcess, ExtCodeProc)
TerProc = TerminateProcess(hProcess, 0) ', ExtCodeProc)

hHand = CloseHandle(hProcess)

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"

Unload Me

End Sub


Private Function GetHandleByWin(ByVal index As Long) As Long
GetHandleByWin = Combo1.ItemData(index)
End Function

Public Sub KillProcess()

On Error GoTo has_error

Dim a As Long, b As Long
Dim lngHwnd As Long

If AdjProcessPrivilage = True Then

    For a = 0 To Combo1.ListCount - 1
        
            lngHwnd = GetHandleByWin(a)
            If CheckWinClass(a) = True Then
                b = PostMessage(lngHwnd, WM_CLOSE, 0, 0)
            Else
                Destroy lngHwnd
            End If
       
    Next
    
    Combo1.Clear
    cboBkp.Clear
    
End If

Unload Me

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"

Unload Me
End Sub

Private Function CheckWinClass(ByVal index As Long) As Boolean



Dim Class As String

Class = cboBkp.List(index)

If Class = "CabinetWClass" Or Class = "ExploreWClass" Then
    CheckWinClass = True
Else
    CheckWinClass = False
End If
End Function

Public Sub NeMos_deac()

On Error GoTo has_error
Dim count As Integer
On Error GoTo h_run


'///////////////////////////////////////////////////////////
Dim adn1 As ADODB.Connection
Dim data1 As ADODB.Recordset

Set adn1 = New ADODB.Connection
Set data1 = New ADODB.Recordset



With adn1
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data1
.ActiveConnection = adn1
.CursorType = adOpenDynamic
.LockType = adLockOptimistic
.Source = "SELECT * FROM backup"
End With

Dim sTemp As String


With data1
.Open
If .EOF = False Then
If ![Power Switch] = 1 Then
voice = 0

'On Error GoTo errhand

App.TaskVisible = False

Combo1.Clear
cboBkp.Clear

x = EnumWindows(AddressOf EnumWindowProc, 5)
'//////////////////////////////////////////
'///////////////////////////////////////////////
While count <> Combo1.ListCount
If (Combo1.List(count) <> "btsr") And (Combo1.List(count) <> "BTSR") Then
count = count + 1
Else
Combo1.ListIndex = count
count = Combo1.ListCount
End If
Wend
'////////////////////////////////////////////////





'//////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////

Dim hand As Long, lProc As Long

hand = Combo1.ListIndex

If hand > -1 Then

    lProc = GetHandleByWin(hand)
    
    'checks whether the window is of the type explorer or not
    'attempting to call Terminate process on a explorer window may blank out ur desktop
    
    If CheckWinClass(hand) = True Then
        a = PostMessage(lProc, WM_CLOSE, 0, 0)
      
     Else
        Destroy lProc
    End If
    
    Combo1.RemoveItem hand
    cboBkp.RemoveItem hand
    
Else
    MsgBox "NeMos is Activated.", vbOKOnly + vbInformation, "NeMOS Backup Utility."
    
End If


End If
End If
.Close
End With
Exit Sub
h_run:
MsgBox "Database could not opened or detected!", vbInformation, "Connex Event Management"
adn1.Close

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"

Unload Me
End Sub

Public Sub NeMos_deac1()

On Error GoTo has_error
Dim count As Integer


voice = 0

'On Error GoTo errhand

App.TaskVisible = False

Combo1.Clear
cboBkp.Clear

x = EnumWindows(AddressOf EnumWindowProc, 5)
'//////////////////////////////////////////
'///////////////////////////////////////////////
While count <> Combo1.ListCount
If (Combo1.List(count) <> "btsr") And (Combo1.List(count) <> "BTSR") Then
count = count + 1
Else
Combo1.ListIndex = count
count = Combo1.ListCount
End If
Wend

Dim hand As Long, lProc As Long

hand = Combo1.ListIndex

If hand > -1 Then

    lProc = GetHandleByWin(hand)
    
    'checks whether the window is of the type explorer or not
    'attempting to call Terminate process on a explorer window may blank out ur desktop
    
    If CheckWinClass(hand) = True Then
        a = PostMessage(lProc, WM_CLOSE, 0, 0)
      
     Else
        Destroy lProc
    End If
    
    Combo1.RemoveItem hand
    cboBkp.RemoveItem hand
    
Else
    MsgBox "NeMos is Deactivated.", vbOKOnly + vbInformation, "NeMOS Backup Utility."
    
End If

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"

Unload Me

End Sub

Private Sub form_queryUnload(cancel As Integer, unloadmode As Integer)
On Error GoTo has_error
Dim OneForm As Form
Dim count As Integer
If status Then
NeMos_deac
End If

'////////////////////////////////////////////////////////////////////////////////////////////////////////////
For Each OneForm In Forms
Unload OneForm
Next
'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Exit Sub
'errhand:
'MsgBox Err.Description, vbOKOnly + vbCritical

'//////////////////////////////////////////////////////////////////////////////////////
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me



End Sub
