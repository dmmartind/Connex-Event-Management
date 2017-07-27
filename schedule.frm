VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule"
   ClientHeight    =   6045
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11670
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox DBCombo1 
      Height          =   315
      Left            =   4560
      TabIndex        =   33
      Text            =   "DBCombo1"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox total1 
      Height          =   285
      Left            =   5160
      TabIndex        =   32
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSMask.MaskEdBox time_out1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   255
      Left            =   9360
      TabIndex        =   31
      Top             =   2880
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "hh:mm AM/PM"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox time_in1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   30
      Top             =   2880
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "hh:mm AM/PM"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   255
      Left            =   1440
      TabIndex        =   28
      Top             =   2400
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "M/d/yyyy"
      Format          =   82640899
      CurrentDate     =   37409
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "schedule.frx":0000
      TabIndex        =   27
      Top             =   2400
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   255
      Left            =   1440
      TabIndex        =   26
      Top             =   1920
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "M/d/yyy"
      Format          =   82640899
      CurrentDate     =   37409
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   10920
      OleObjectBlob   =   "schedule.frx":0074
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   7800
      OleObjectBlob   =   "schedule.frx":2A65D
      TabIndex        =   25
      Top             =   2400
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   7800
      OleObjectBlob   =   "schedule.frx":2A6D1
      TabIndex        =   24
      Top             =   1920
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   7920
      OleObjectBlob   =   "schedule.frx":2A745
      TabIndex        =   23
      Top             =   1440
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   7800
      OleObjectBlob   =   "schedule.frx":2A7B1
      TabIndex        =   22
      Top             =   960
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   7920
      OleObjectBlob   =   "schedule.frx":2A825
      TabIndex        =   21
      Top             =   2880
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "schedule.frx":2A893
      TabIndex        =   20
      Top             =   2880
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "schedule.frx":2A8FF
      TabIndex        =   19
      Top             =   1920
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "schedule.frx":2A979
      TabIndex        =   18
      Top             =   1440
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "schedule.frx":2A9E3
      TabIndex        =   17
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   2520
      Picture         =   "schedule.frx":2AA4B
      ScaleHeight     =   435
      ScaleWidth      =   6795
      TabIndex        =   16
      Top             =   120
      Width           =   6855
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   11415
      Begin VB.CommandButton cmdnext 
         Caption         =   "Next"
         Height          =   255
         Left            =   2400
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   255
         Left            =   10200
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   8040
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete All"
         Height          =   255
         Left            =   4920
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   873
         _Version        =   393216
         AllowUpdate     =   -1  'True
         DefColWidth     =   240
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   2
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Button          =   -1  'True
               ColumnWidth     =   4995.213
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   4995.213
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   4995.213
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   3000.189
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox datetime 
      DataField       =   "Date_Time"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox phone1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9360
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox event_admin1 
      DataField       =   "Administrator"
      Height          =   315
      Left            =   9360
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Reset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   9600
      TabIndex        =   12
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   5520
      Width           =   1935
   End
   Begin VB.ComboBox hour_type1 
      DataField       =   "Hour_Type"
      Height          =   315
      Left            =   9360
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox room1 
      DataField       =   "Room"
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox group1 
      DataField       =   "Name/Group"
      Enabled         =   0   'False
      Height          =   285
      Left            =   9360
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox event1 
      DataField       =   "Event"
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin VB.Menu menu1 
      Caption         =   "Main"
      Begin VB.Menu E_xit 
         Caption         =   "Exit"
         Index           =   1
      End
   End
   Begin VB.Menu record_control 
      Caption         =   "Record Control"
      Begin VB.Menu new_record 
         Caption         =   "New Record"
      End
      Begin VB.Menu edit_record 
         Caption         =   "Edit Record"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu search_record 
         Caption         =   "Search Record"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Dim record_control2 As Integer
Dim record_control3 As Integer
Dim opps As Integer
Dim adn20 As ADODB.Connection
Attribute adn20.VB_VarHelpID = -1
Dim WithEvents data2 As ADODB.Recordset
Attribute data2.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim record_in As Integer
Dim grep As String
Dim do_add As Boolean







Private Sub Add_Click()

On Error GoTo has_error

Dim connection_string As String

' /////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
connection_string = Form1.connection_string
'data2.Close
'adn20.Close

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
.Source = "Select * From Schedule_Info Order by [Event] ASC"
End With


Rem Set DataGrid1.DataSource = data2
Rem DataGrid1.ReBind











Dim drive As Integer
drive = 0
check_form drive
If drive = 0 Then

'//////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////
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


hour1 = Hour(time_in1.Text)
hour2 = Hour(time_out1.Text)

min1 = Minute(time_in1.Text)
min2 = Minute(time_out1.Text)

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
total1_s = count_hc & ":" & count_mc
total1.Text = total1_s


If drive = 0 Then

If (record_control2 = 0) Then
datetime.Text = Now()
With data1
.Open
.AddNew
.Fields("Event") = event1.Text
.Fields("Administrator") = event_admin1.Text
.Fields("Room") = room1.Text
.Fields("Hour_Type") = hour_type1.Text
.Fields("Begin Date") = date1.Value
.Fields("End Date") = date2.Value
.Fields("Time") = time_in1.Text
.Fields("Date_Time") = datetime.Text
.Fields("Time_Out") = time_out1.Text
.Fields("Name/Group") = group1.Text
.Fields("total_hours") = total1.Text
.Fields("phone#") = phone1.Text
.Update
.Close
End With

record_control2 = -1
record_control3 = -1



Add.Enabled = False
Reset.Enabled = False
event1.Enabled = False
room1.Enabled = False
date1.Enabled = False
date2.Enabled = False
time_in1.Enabled = False
hour_type1.Enabled = False
group1.Enabled = False
time_out1.Enabled = False
event_admin1.Enabled = False
DataGrid1.Enabled = False
menu1.Enabled = True


opps = -1
End If
End If

If record_control2 = 1 Then
datetime.Text = grep


With data1
.Open
.Find "[Event]= '" & event1.Text & "'"
.Delete
.AddNew
.Fields("Event") = event1.Text
.Fields("Administrator") = event_admin1.Text
.Fields("Room") = room1.Text
.Fields("Hour_Type") = hour_type1.Text
.Fields("Begin Date") = date1.Value
.Fields("End Date") = date2.Value
.Fields("Time") = time_in1.Text
.Fields("Date_Time") = datetime.Text
.Fields("Time_Out") = time_out1.Text
.Fields("Name/Group") = group1.Text
.Fields("total_hours") = total1.Text
.Update
.Close
End With





record_control2 = -1
record_control3 = -1


Add.Enabled = False
Reset.Enabled = False
event1.Enabled = False
room1.Enabled = False
date1.Enabled = False
date2.Enabled = False
time_in1.Enabled = False
new_record.Enabled = True
edit_record.Enabled = False
search_record.Enabled = True
hour_type1.Enabled = False
group1.Enabled = False
time_out1.Enabled = False
event_admin1.Enabled = False
DataGrid1.Enabled = False
menu1.Enabled = True
cmdAdd.Enabled = False
cmdnext.Enabled = False
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
cmdRefresh.Enabled = False






opps = -1
End If
End If

DBCombo1.Enabled = False
DBCombo1.Visible = False
Reset.Enabled = True
adn.Close

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub DataGrid1_Click()

On Error GoTo has_error



DataGrid1.ReBind
DataGrid1.Columns(1).Button = True
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(4).Visible = False
DataGrid1.Columns(0) = event1.Text
DataGrid1.Columns(4) = date1.Value
If DataGrid1.Columns(1) <> "" And DataGrid1.Columns(2) = "" Then
    DataGrid1.Col = 2
    DataGrid1.Columns(2) = ""
    DataGrid1.SetFocus
End If
If ((DataGrid1.Columns(1) <> "") And (DataGrid1.Columns(2) <> "") And (DataGrid1.Columns(3) = "")) Then
    DataGrid1.Col = 3
    DataGrid1.Columns(3) = ""
    DataGrid1.SetFocus
End If
DataGrid1.SetFocus
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub



Private Sub DBCombo1_Click()
On Error GoTo has_error
DataGrid1.Columns(1) = DBCombo1.Text
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub E_xit_Click(index As Integer)
On Error GoTo has_error
Unload Me
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub event_admin1_Click()
On Error GoTo has_error
Dim connection_string As String

'////////////////////////////////////////////////
connection_string = Form1.connection_string

Dim adn4 As ADODB.Connection
Dim data5 As ADODB.Recordset

Set adn4 = New ADODB.Connection
Set data5 = New ADODB.Recordset

With adn4
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data5
.ActiveConnection = adn4
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Customers Order by [Name] ASC"
End With
'///////////////////////////////////////////////
With data5
.Open
.Find "[name]='" & event_admin1.Text & "'"
If .EOF = False Then
phone1.Text = ![phone]
group1.Text = ![Group]
End If
.Close
End With

adn4.Close
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub



Private Sub Reset_Click()

On Error GoTo has_error
record_control2 = -1
record_control3 = -1
Dim connection_string As String

' /////////////////////////////////////////////////////////////////////////////
' ////////////////////////////////////////////////////////////////////////////
connection_string = Form1.connection_string


Dim adn As ADODB.Connection
Dim data1 As ADODB.Recordset

Set adn = New ADODB.Connection
Set data1 = New ADODB.Recordset

Dim adn20 As ADODB.Connection
Dim data2 As ADODB.Recordset

Set adn20 = New ADODB.Connection
Set data2 = New ADODB.Recordset

With adn
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With adn20
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data1
.ActiveConnection = adn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Schedule_Info Order by [Event] ASC"
End With


With data2
.ActiveConnection = adn20
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From activity2 where [Event]= '" & event1.Text & "' AND [Event_Date]= #" & date1.Value & "# Order by [Event] DESC"
End With



 With data2
  .Open
 Set DataGrid1.DataSource = data2
DataGrid1.ReBind

If record_in = 1 Then

 Rem On Error GoTo DeleteErr
  
    While ((.EOF = False))
  .Delete
  .MoveLast
  Wend
  'DataGrid1.Refresh
  
  
DataGrid1.Enabled = False

End If
End With
record_in = 0


event1.Enabled = False
room1.Enabled = False
date1.Enabled = False
date2.Enabled = False
time_in1.Enabled = False
Add.Enabled = False
Delete.Enabled = False
Reset.Enabled = False
edit_record.Enabled = False
new_record.Enabled = True
search_record.Enabled = True
hour_type1.Enabled = False
group1.Enabled = False
time_out1.Enabled = False
event_admin1.Enabled = False
DataGrid1.Enabled = False
cmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
cmdRefresh.Enabled = False
cmdnext.Enabled = False
menu1.Enabled = True

 


  
  

opps = -1
DataGrid1.Enabled = False
DBCombo1.Enabled = False


adn.Close
data2.Close
adn20.Close
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub Delete_Click()
On Error GoTo has_error
'///////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////
cmdDelete_Click
Dim connection_string As String

connection_string = Form1.connection_string


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
.Source = "Select * From Schedule_Info where [Event]= '" & event1.Text & "'  AND [Event_Date]= #" & date1.Value & "# Order by [Event] DESC"
.Open
End With

'.Source = "Select * From activity2 where [Event]= '" & Combo1.Text & "' AND [Event_Date]= #" & temp1 & "#"
'/////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////
With data1
'.Find "[Event]= '" & event1.Text & "'"
.Delete
If .EOF Then .MoveLast
End With
record_control2 = -1
Add.Enabled = False
event1.Enabled = False
room1.Enabled = False
date1.Enabled = False
date2.Enabled = False
time_in1.Enabled = False

hour_type1.Enabled = False
hour_type1.Enabled = False
group1.Enabled = False
time_out1.Enabled = False
event_admin1.Enabled = False
new_record.Enabled = True
edit_record.Enabled = False
search_record.Enabled = True
Delete.Enabled = False
record_control3 = -1

adn.Close

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub edit_record_Click()
On Error GoTo has_error
record_in = 1
record_control2 = 1
record_control3 = -1
Add.Enabled = True
Delete.Enabled = False
event1.Enabled = True
room1.Enabled = True
date1.Enabled = True
date2.Enabled = True
time_in1.Enabled = True
hour_type1.Enabled = True
group1.Enabled = False
time_out1.Enabled = True
event_admin1.Enabled = True
Reset.Enabled = True
cmdAdd.Enabled = True
cmdDelete.Enabled = True
cmdRefresh.Enabled = True
cmdUpdate.Enabled = True
opps = 1
grep = datetime.Text
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
On Error GoTo has_error
Skin1.ApplySkin Me.hwnd
Dim connection_string As String

connection_string = Form1.connection_string
record_control2 = -1
record_control3 = -1
'record_control4 = -1


Dim adn1 As ADODB.Connection
Dim adn2 As ADODB.Connection
Dim adn3 As ADODB.Connection
Dim adn4 As ADODB.Connection
Dim adn5 As ADODB.Connection

Dim data1 As ADODB.Recordset
Dim data3 As ADODB.Recordset
Dim data4 As ADODB.Recordset
Dim data5 As ADODB.Recordset
Dim Data6 As ADODB.Recordset

Set adn1 = New ADODB.Connection
Set adn2 = New ADODB.Connection
Set adn3 = New ADODB.Connection
Set adn4 = New ADODB.Connection
Set adn5 = New ADODB.Connection


Set data1 = New ADODB.Recordset
Set data3 = New ADODB.Recordset
Set data4 = New ADODB.Recordset
Set data5 = New ADODB.Recordset
Set Data6 = New ADODB.Recordset

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


With adn3
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With adn4
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With adn5
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data1
.ActiveConnection = adn1
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Employee Order by [Employee_Name]"
End With


With data3
.ActiveConnection = adn2
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Event_Room Order by [Room_#] ASC"
End With

With data4
.ActiveConnection = adn3
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Event_Time"
End With

With data5
.ActiveConnection = adn4
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Customers Order by [Name]"
End With

With Data6
.ActiveConnection = adn5
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From activity_type"
End With


event1.Enabled = False
room1.Enabled = False
date1.Enabled = False
date2.Enabled = False
time_in1.Enabled = False
Add.Enabled = False
Delete.Enabled = False
Reset.Enabled = False
edit_record.Enabled = False
hour_type1.Enabled = False
group1.Enabled = False
time_out1.Enabled = False
event_admin1.Enabled = False
opps = -1
cmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
cmdRefresh.Enabled = False
cmdnext.Enabled = False
DataGrid1.Enabled = False

data3.Open
With data3


Do Until data3.EOF
    If ![Room_#] <> "" Then
        room1.AddItem ![Room_#]
    End If
        .MoveNext
Loop
.Close
End With
Rem //////////////////////////////
data4.Open
With data4


Do Until data4.EOF
    If ![Event_Time] <> "" Then
        hour_type1.AddItem ![Event_Time]
    End If
        .MoveNext
Loop
.Close
End With
Rem ////////////////////////////////

With data5
.Open

Do Until .EOF
    If ![name] <> "" Then
        event_admin1.AddItem ![name]
    End If
        .MoveNext
Loop
.Close
End With

With Data6
.Open
While .EOF = False
DBCombo1.AddItem ![activity]
.MoveNext
Wend
.Close
End With
DBCombo1.Visible = False


adn1.Close
adn2.Close
adn3.Close
adn4.Close
adn5.Close
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me



End Sub

Private Sub new_record_Click()

On Error GoTo has_error
record_in = 1
record_control2 = 0
record_control3 = 1
Dim connection_string As String

'//////////////////////////////////////////////////////////
connection_string = Form1.connection_string

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
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Employee Order by [Employee_Name]"
End With

With data1
.Open
.AddNew
End With


'///////////////////////////////////////////////////////////////

event1.Enabled = True
event1.Text = ""
room1.Enabled = True
room1.Text = ""
date1.Enabled = True
date2.Enabled = True
new_record.Enabled = False
edit_record.Enabled = False
search_record.Enabled = False
time_in1.Enabled = True
time_in1.Text = ""
hour_type1.Enabled = True
hour_type1.Text = ""
group1.Enabled = False
group1.Text = ""
time_out1.Enabled = True
time_out1.Text = ""
event_admin1.Enabled = True
event_admin1.Text = ""


Add.Enabled = False
Reset.Enabled = False
Delete.Enabled = False
cmdAdd.Enabled = True
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
cmdRefresh.Enabled = False
DataGrid1.Enabled = True
opps = 1
DataGrid1.ClearFields
DataGrid1.Enabled = False
menu1.Enabled = False
Reset.Enabled = True

connection_string = Form1.connection_string

Set adn20 = New ADODB.Connection
Set data2 = New ADODB.Recordset

With adn20
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data2
.ActiveConnection = adn20
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From activity2 where [Event]= '" & event1.Text & "'Order by [Event]"
.Open
End With


Set DataGrid1.DataSource = data2


DataGrid1.ReBind
DataGrid1.Columns(1).Button = True
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(4).Visible = False

adn1.Close

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub



Private Sub search_record_Click()

On Error GoTo has_error
Search_e4.Show vbModal
Dim connection_string As String



record_control3 = -1

'/////////////////////////////////////////////////////////////////////////////
connection_string = Form1.connection_string

Set adn20 = New ADODB.Connection
Set data2 = New ADODB.Recordset

With adn20
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data2
.ActiveConnection = adn20
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From activity2 where [Event]= '" & event1.Text & "' Order by [Event] ASC"
.Open
End With
DataGrid1.ClearFields
Set DataGrid1.DataSource = data2
DataGrid1.ReBind
DataGrid1.Columns(1).Button = True
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(4).Visible = False
'//////////////////////////////////////////////////////////////////////////////////
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me


End Sub


Private Sub form_queryUnload(cancel As Integer, unloadmode As Integer)

On Error GoTo has_error


Dim temp As Integer
Dim connection_string As String

'////////////////////////////////////////////////////////////////////
connection_string = Form1.connection_string
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
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Employee Order by [Employee_Name]"
End With

'/////////////////////////////////////////////////////////////////////
Dim adn20 As ADODB.Connection
Dim data2 As ADODB.Recordset

Set adn20 = New ADODB.Connection
Set data2 = New ADODB.Recordset

With adn20
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data2
.ActiveConnection = adn20
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From activity2"
.Open
End With
Set DataGrid1.DataSource = data2
DataGrid1.ReBind
data2.Close
adn20.Close

'/////////////////////////////////////////////////////////////////////










data1.Open
If opps = 1 Then
temp = MsgBox("Do you really want to quit?", vbYesNo, "Connex Event Management")
Select Case temp
Case 6
    data1.cancel
    Unload Me
Case 7
    cancel = 1
End Select
Else
data1.cancel
Unload Me
End If
data1.Close
adn1.Close

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub check_form(ByRef drive As Integer)
'////////////////////////////////////////////////////////////////////


On Error GoTo has_error
Dim connection_string As String


connection_string = Form1.connection_string
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
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Employee Order by [Employee_Name]"
End With

'/////////////////////////////////////////////////////////////////////
drive = 0
If event1.Text = "" Then
    MsgBox "Please enter the event name", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf room1.Text = "" Then
  MsgBox "Please enter the room number", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf time_in1.Text = "" Then
  MsgBox "Please enter the begin time the event", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf time_out1.Text = "" Then
  MsgBox "Please enter the ending date of the the event", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf event_admin1.Text = "" Then
  MsgBox "Please enter the administrator of the event", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf hour_type1.Text = "" Then
  MsgBox "Please enter the type of hours of the event", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
End If

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub



Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  
  On Error GoTo has_error
  
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
  
  Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub cmdAdd_Click()

On Error GoTo has_error
Dim connection_string As String

connection_string = Form1.connection_string

Set adn20 = New ADODB.Connection
Set data2 = New ADODB.Recordset

With adn20
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data2
.ActiveConnection = adn20
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From activity2 where [Event]= '" & event1.Text & "'Order by [Event]"
.Open
End With

Set DataGrid1.DataSource = data2
DataGrid1.ReBind

DataGrid1.Enabled = True
DataGrid1.Columns(1).Width = 3750
DataGrid1.Columns(1).AllowSizing = False
DataGrid1.Columns(2).Width = 3750
DataGrid1.Columns(2).AllowSizing = False
DataGrid1.Columns(3).Width = 3750
DataGrid1.Columns(3).AllowSizing = False
DataGrid1.Columns(1).Button = True
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(0).AllowSizing = False
DataGrid1.Columns(4).Visible = False
DataGrid1.Columns(4).AllowSizing = False


Rem   On Error GoTo AddErr
  data2.AddNew

mbDataChanged = False
Rem DataGrid1.SetFocus
Rem   SendKeys "{down}"
cmdAdd.Enabled = False
cmdnext.Enabled = True
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
cmdRefresh.Enabled = False
     
      DBCombo1.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight + 3350
      DBCombo1.Left = DataGrid1.Left + DataGrid1.Columns(ColIndex).Left + 450
      ' Width and Height properties can be set a design time
      ' The width of the list does not have to be the same as the width of the grid column
      DBCombo1.Width = DataGrid1.Columns(1).Width - 150
      '//DBCombo1.Height = 225
      DBCombo1.Visible = True
      If DBCombo1.Visible = True Then
         DBCombo1.Text = DataGrid1.Text
         DBCombo1.ZOrder ' make sure the list is on top of the grid
         DataGrid1.SetFocus
         DBCombo1.Enabled = True
         End If

  Exit Sub
Rem AddErr:
Rem   MsgBox Err.Description
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub cmdDelete_Click()

On Error GoTo has_error

  

 Rem On Error GoTo DeleteErr
  With data2
  While ((.EOF = False))
  .Delete
  .MoveNext
  Wend
  
  DataGrid1.Refresh
  End With
  cmdAdd.Enabled = True
  cmdnext.Enabled = False
  cmdDelete.Enabled = False
  cmdUpdate.Enabled = False
  cmdRefresh.Enabled = True
   Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
  
  End Sub
Rem DeleteErr:
Rem MsgBox Err.Description
Rem End Sub

Private Sub cmdRefresh_Click()

On Error GoTo has_error

Set DataGrid1.DataSource = data2
DataGrid1.ReBind
DataGrid1.Columns(1).Button = True
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(4).Visible = False

 
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  data2.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
 Exit Sub
 
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
  
End Sub

Private Sub cmdUpdate_Click()

On Error GoTo has_error
  
Set DataGrid1.DataSource = data2
DataGrid1.ReBind
DataGrid1.Columns(1).Button = True
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(4).Visible = False
cmdAdd.Enabled = False
cmdnext.Enabled = False
cmdDelete.Enabled = False
cmdRefresh.Enabled = False
cmdUpdate.Enabled = False
Add.Enabled = True
Reset.Enabled = True
DataGrid1.Enabled = False


  On Error GoTo UpdateErr
  data2.UpdateBatch adAffectAll
  
  Exit Sub
UpdateErr:
  MsgBox Err.Description
  Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub cmdClose_Click()

On Error GoTo has_error
  Unload Me
 Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo has_error
  
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
     ' cmdLast_Click
    Case vbKeyHome
     ' cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
      '  cmdFirst_Click
      Else
        'cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
       ' cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
  
  Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub Form_Unload(cancel As Integer)
On Error GoTo has_error
  Screen.MousePointer = vbDefault
  Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub cmdNext_Click()

On Error GoTo has_error
  '//On Error GoTo GoNextError


  If Not data2.EOF Then data2.MoveNext
  If data2.EOF And data2.RecordCount > 0 Then
    data2.MoveLast
  End If
  mbDataChanged = False
cmdAdd.Enabled = True
cmdnext.Enabled = False
cmdDelete.Enabled = True
cmdUpdate.Enabled = True
cmdRefresh.Enabled = False

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

