VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Clock-Out"
   ClientHeight    =   4650
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7080
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form4.frx":0000
      TabIndex        =   15
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox total1 
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   120
      Picture         =   "Form4.frx":0072
      ScaleHeight     =   435
      ScaleWidth      =   6675
      TabIndex        =   12
      Top             =   0
      Width           =   6735
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   1680
      ScaleHeight     =   1995
      ScaleWidth      =   3675
      TabIndex        =   11
      Top             =   2520
      Width           =   3735
      Begin VB.ListBox List1 
         Height          =   2010
         ItemData        =   "Form4.frx":312E
         Left            =   0
         List            =   "Form4.frx":3130
         MultiSelect     =   2  'Extended
         TabIndex        =   4
         Top             =   0
         Width           =   3735
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3960
      OleObjectBlob   =   "Form4.frx":3132
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   4800
      OleObjectBlob   =   "Form4.frx":2D71B
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   4800
      OleObjectBlob   =   "Form4.frx":2D781
      TabIndex        =   9
      Top             =   720
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form4.frx":2D7E7
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "Form4.frx":2D84F
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox time1 
      DataField       =   "Date"
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox name1 
      DataField       =   "Employee_Name"
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.ComboBox event1 
      DataField       =   "Event_ID"
      Height          =   315
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Clock_Out 
      Caption         =   "Clock-Out"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox date1 
      DataField       =   "Out_Time"
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin MSMask.MaskEdBox event_date1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   1680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Menu file_1 
      Caption         =   "File"
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Dim record_control2 As Integer
Dim opps As Integer


Private Sub Reset_Click()
On Error GoTo has_error
name1.Locked = True
name1.Text = Form1.user_name1
date1.Text = Date
time1.Text = Time
data1.Refresh
data2.Refresh

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub
Private Sub Clock_Out_Click()

On Error GoTo has_error
'************************************************************************************8
'***********************************************************************************
Dim connection_string As String
connection_string = Form1.connection_string
Dim flag As Boolean
flag = False

Dim adn As ADODB.Connection
Dim data1 As ADODB.Recordset

Set adn = New ADODB.Connection
Set data1 = New ADODB.Recordset

With adn
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data1
.ActiveConnection = adn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Check_in Where [Event_ID]= '" & event1.Text & "' and [Event_Date]= # " & event_date1.Text & " # Order by [Event_ID] DESC"
.Open
End With


'////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////
'***********************
Dim hour1 As Integer
Dim min1 As Integer
Dim hour2 As Integer
Dim min2 As Integer
Dim Dir1 As Integer
Dim dir2 As Integer
Dim total_h As Integer
Dim total_m As Integer
Dim count_h As Integer
Dim count_m As Integer
Dim count_hc As Integer
Dim count_mc As Integer
'***************************


If data1.EOF = False Then

hour1 = Hour(data1!In_Time)
hour2 = Hour(time1.Text)

min1 = Minute(data1!In_Time)
min2 = Minute(time1.Text)
With data1
.Close
End With

count_h = hour1
count_m = min1
count_hc = 0
count_mc = 0

'///////////////hours count//////////
While (count_h < hour2)
count_hc = count_hc + 1
If count_h > 23 Then
count_h = count_h - 12
End If
count_h = count_h + 1
Wend

While (count_h > hour2)
count_hc = count_hc + 1
If count_h < 23 Then
count_h = count_h - 12
End If
count_h = count_h + 1
If count_h > 12 Then
count_h = count_h - 12
End If
Wend
'//////////////hours count////////////

'/////////////minutes count///////////
While (count_m < min2)
count_mc = count_mc + 1
If count_m > 60 Then
count_m = count_m - 60
count_h = count_h + 1
End If
count_m = count_m + 1
Wend
'/////////////minutes count///////////
Dim total1_s As String

If count_mc < 10 Then
total1_s = count_hc & ":" & "0" & count_mc
Else
total1_s = count_hc & ":" & count_mc
End If

total1.Text = total1_s
'//////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////
'***********************************************************************
'************************************************************************
If (record_control2 = 0) Then
With data1
.Open
.Fields(5) = time1.Text
Dim setup As Integer

Dim txt As String
setup = 0

'///////////////////////////////////////////////////////////////////////
While (setup < List1.ListCount)
If List1.Selected(setup) = True Then
txt = txt & setup & ","
End If
setup = setup + 1
Wend
'///////////////////////////////////////////////////////////////////////



.Fields(6) = txt
.Fields(8) = total1.Text
.Update
.Close
adn.Close
End With
End If
Else
data1.Close
adn.Close
End If
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub


Private Sub event1_Click()

On Error GoTo has_error
'**********************************************************
List1.Clear
Dim connection_string As String
connection_string = Form1.connection_string
name1.Locked = True
name1.Text = Form1.user_name1
date1.Text = Date
time1.Text = Time

Dim adn As ADODB.Connection
Dim adn1 As ADODB.Connection
Dim adn2 As ADODB.Connection

Dim data1 As ADODB.Recordset
Dim data2 As ADODB.Recordset
Dim data3 As ADODB.Recordset

Set adn = New ADODB.Connection
Set adn1 = New ADODB.Connection
Set adn2 = New ADODB.Connection

Set data1 = New ADODB.Recordset
Set data2 = New ADODB.Recordset
Set data3 = New ADODB.Recordset

With adn
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With adn1
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With adn2
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With


With data1
.ActiveConnection = adn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Schedule_Info where [Event]= '" & event1.Text & "' Order by [Event]"
End With


With data1
.Open
If .EOF = False Then
event_date1.Text = ![Begin Date]
End If
.Close
End With
adn.Close


With data2
.ActiveConnection = adn1
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
'.Source = "Select * From activity2 Where [Event]= '" & event1.Text & "' and Where [Event_Date]= '" & event_date1.Text & "' Order by [Event] DESC"
'.Source = "Select * From activity2 Where [Event]= '" & List1.List(count) & "' AND [Event_Date] >= '#" & begin.Text & "#'"
.Source = "Select * From activity2 Where [Event]= '" & event1.Text & "' AND [Event_Date] = #" & event_date1.Text & "#"
End With
'**********************************************************

'/////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////




With data2
.Open
'.Find "[Event_Date]= '" & event_date1.Text & "'"
If .EOF = False Then
event_date1.Text = ![event_date]
While .EOF = False
List1.AddItem ![Equipment]
.MoveNext
Wend
End If
.Close
End With

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub exit_Click()
On Error GoTo has_error

Unload Me

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub Form_Load()
'On Error GoTo has_error
Skin1.ApplySkin Me.hwnd
Dim connection_string As String
connection_string = Form1.connection_string


name1.Locked = True
name1.Text = Form1.user_name1
date1.Text = Date
time1.Text = Time

Dim adn As ADODB.Connection
Dim adn1 As ADODB.Connection
Dim adn2 As ADODB.Connection

Dim data1 As ADODB.Recordset
Dim data2 As ADODB.Recordset
Dim data3 As ADODB.Recordset

Set adn = New ADODB.Connection
Set adn1 = New ADODB.Connection
Set adn2 = New ADODB.Connection

Set data1 = New ADODB.Recordset
Set data2 = New ADODB.Recordset
Set data3 = New ADODB.Recordset

With adn
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With adn1
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data1
.ActiveConnection = adn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Schedule_Info Order by [Event]"
End With

With data2
.ActiveConnection = adn1
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From activity2 where 1=2"
End With





With data1
.Open
Do Until data1.EOF
    If ![Event] <> "" Then
        event1.AddItem ![Event]
    End If
        .MoveNext
Loop
.Close
End With

name1.Text = Form1.user_name1

Exit Sub
'has_error:
'MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
'Unload Me

End Sub
Private Sub form_queryUnload(cancel As Integer, unloadmode As Integer)
On Error GoTo has_error
'*********************************************************************
Dim connection_string As String
connection_string = Form1.connection_string
name1.Locked = True
name1.Text = Form1.user_name1
date1.Text = Date
time1.Text = Time

Dim adn As ADODB.Connection
Dim adn1 As ADODB.Connection
Dim adn2 As ADODB.Connection

Dim data1 As ADODB.Recordset
Dim data2 As ADODB.Recordset
Dim data3 As ADODB.Recordset

Set adn = New ADODB.Connection
Set adn1 = New ADODB.Connection
Set adn2 = New ADODB.Connection

Set data1 = New ADODB.Recordset
Set data2 = New ADODB.Recordset
Set data3 = New ADODB.Recordset

With adn
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With adn1
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With adn2
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With


With data1
.ActiveConnection = adn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Check_In Order by [Event_ID]"
End With
'**********************************************************************
With data1
Dim temp As Integer
If opps = 1 Then
temp = MsgBox("Do you really want to quit?", vbYesNo, "Connex Event Management")
Select Case temp
Case 6
    .cancel
    Unload Me
Case 7
    cancel = 1
End Select
Else
.cancel

Unload Me
End If
End With

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub






