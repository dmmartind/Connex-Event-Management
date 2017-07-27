VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form20 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6165
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Time_Date 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   24
      Top             =   3360
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   255
      Left            =   5520
      OleObjectBlob   =   "schedule_view.frx":0000
      TabIndex        =   23
      Top             =   2880
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5640
      OleObjectBlob   =   "schedule_view.frx":008C
      Top             =   4920
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   4440
      OleObjectBlob   =   "schedule_view.frx":2A675
      TabIndex        =   18
      Top             =   2400
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   4320
      OleObjectBlob   =   "schedule_view.frx":2A6E9
      TabIndex        =   17
      Top             =   1920
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   4440
      OleObjectBlob   =   "schedule_view.frx":2A75D
      TabIndex        =   16
      Top             =   1440
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   4320
      OleObjectBlob   =   "schedule_view.frx":2A7C9
      TabIndex        =   15
      Top             =   960
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "schedule_view.frx":2A83D
      TabIndex        =   14
      Top             =   3480
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "schedule_view.frx":2A8AB
      TabIndex        =   13
      Top             =   3000
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "schedule_view.frx":2A917
      TabIndex        =   12
      Top             =   1440
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "schedule_view.frx":2A981
      TabIndex        =   11
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   480
      Picture         =   "schedule_view.frx":2A9E9
      ScaleHeight     =   435
      ScaleWidth      =   6795
      TabIndex        =   10
      Top             =   120
      Width           =   6855
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   8775
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2055
         Left            =   0
         TabIndex        =   25
         Top             =   120
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   3625
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
               Format          =   ""
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
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   2
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox datetime 
      DataField       =   "Date_Time"
      DataSource      =   "data1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox phone1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox event_admin1 
      DataField       =   "Administrator"
      Height          =   315
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.ComboBox hour_type1 
      DataField       =   "Hour_Type"
      Height          =   315
      Left            =   5760
      Sorted          =   -1  'True
      TabIndex        =   6
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
   Begin VB.TextBox time_in1 
      DataField       =   "Time"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox group1 
      DataField       =   "Name/Group"
      Height          =   285
      Left            =   5760
      TabIndex        =   7
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
   Begin VB.TextBox time_out1 
      DataField       =   "Time_Out"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   3480
      Width           =   2415
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "schedule_view.frx":2D73D
      TabIndex        =   19
      Top             =   1920
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   255
      Left            =   1440
      TabIndex        =   20
      Top             =   1920
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      Format          =   45547521
      CurrentDate     =   37409
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "schedule_view.frx":2D7B7
      TabIndex        =   21
      Top             =   2520
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   2520
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      Format          =   45547521
      CurrentDate     =   37409
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
      Begin VB.Menu search_record 
         Caption         =   "Search Record"
      End
   End
End
Attribute VB_Name = "Form20"
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








Private Sub DataGrid1_Click()
On Error GoTo has_error
Dim connection_string As String

If record_control3 = 1 Then
connection_string = Form1.connection_string
record_control2 = 0

Dim adn As ADODB.Connection
Dim data2 As ADODB.Recordset

Set adn = New ADODB.Connection
Set data2 = New ADODB.Recordset

With adn
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data2
.ActiveConnection = adn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From activity2 Order by [Event] ASC"
.Open
End With
data2.AddNew
DataGrid1.ReBind
DataGrid1.Col = 0
DataGrid1.Text = event1.Text



End If
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub DBGrid1_Click()

On Error GoTo has_error
DBGrid1.Col = 0
DBGrid1.Text = event1.Text
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
data5.Refresh
With data5.Recordset
.Find "[name]='" & event_admin1.Text & "'"
If .EOF = False Then
phone1.Text = ![phone]
group1.Text = ![Group]
End If
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

On Error GoTo has_error
Skin1.ApplySkin Me.hwnd
Dim connection_string As String

connection_string = Form1.connection_string
record_control2 = -1
record_control3 = -1
'record_control4 = -1


Dim adn As ADODB.Connection
Dim adn1 As ADODB.Connection
Dim adn2 As ADODB.Connection
Dim adn3 As ADODB.Connection
Dim adn4 As ADODB.Connection

Dim data2 As ADODB.Recordset
Dim data4 As ADODB.Recordset
Dim data1 As ADODB.Recordset
Dim data3 As ADODB.Recordset
Dim data5 As ADODB.Recordset

Set adn = New ADODB.Connection
Set adn1 = New ADODB.Connection
Set adn2 = New ADODB.Connection
Set adn3 = New ADODB.Connection
Set adn4 = New ADODB.Connection

Set data2 = New ADODB.Recordset
Set data4 = New ADODB.Recordset
Set data3 = New ADODB.Recordset
Set data1 = New ADODB.Recordset
Set data5 = New ADODB.Recordset


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

With data2
.ActiveConnection = adn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From activity2 where 1=2"
.Open
End With

Set DataGrid1.DataSource = data2
DataGrid1.ReBind

With data1
.ActiveConnection = adn1
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Schedule_Info Order by [Event] ASC"
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
.Source = "Select * From Customers Order by [Name] ASC"
End With


event1.Enabled = False
room1.Enabled = False
date1.Enabled = False
date2.Enabled = False
time_in1.Enabled = False
hour_type1.Enabled = False
group1.Enabled = False
time_out1.Enabled = False
event_admin1.Enabled = False
opps = -1
DataGrid1.Enabled = False



With data3
.Open

Do Until data3.EOF
    If ![Room_#] <> "" Then
        room1.AddItem ![Room_#]
    End If
        .MoveNext
Loop
.Close
End With
Rem //////////////////////////////
 
With data4
.Open

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
adn4.Close
adn3.Close
adn2.Close
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub search_record_Click()

On Error GoTo has_error
Dialog.Show vbModal
record_control3 = -1
DataGrid1.Enabled = True
DataGrid1.AllowAddNew = False
DataGrid1.AllowArrows = True
DataGrid1.AllowDelete = False
DataGrid1.AllowUpdate = False
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(4).Visible = False
DataGrid1.Columns(1).Width = 2750
DataGrid1.Columns(2).Width = 2750
DataGrid1.Columns(3).Width = 2750

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub form_queryUnload(cancel As Integer, unloadmode As Integer)


On Error GoTo has_error
Dim temp As Integer
'//////////////////////////////////////////////////////////////////////////////////////
Dim connection_string As String

connection_string = Form1.connection_string



Dim adn As ADODB.Connection
Dim adn1 As ADODB.Connection
Dim data1 As ADODB.Recordset
Dim data2 As ADODB.Recordset


Set adn = New ADODB.Connection
Set adn1 = New ADODB.Connection
Set data1 = New ADODB.Recordset
Set data2 = New ADODB.Recordset



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

'///////////////////////////////////////////////////////////////////////////////////////
With data2
.ActiveConnection = adn1
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From activity2 where 1=2"
.Open
End With

Set DataGrid1.DataSource = data2
DataGrid1.ReBind








With data1
.ActiveConnection = adn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Schedule_Info Order By [Event] ASC"
End With

data1.Open

If opps = 1 Then
temp = MsgBox("Do you really want to quit?", vbYesNo, "Connex Event Management")
Select Case temp
Case 6
    data1.cancel
    data2.cancel
    Unload Me
Case 7
    cancel = 1
End Select
Else
data1.cancel
data2.cancel
Unload Me
End If
adn.Close
adn1.Close

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub check_form(ByRef drive As Integer)

On Error GoTo has_error
drive = 0
If event1.Text = "" Then
    MsgBox "Please enter the event name", vbCritical, "Connex Event Management"
    data1.Recordset.cancel
    drive = 1
ElseIf room1.Text = "" Then
  MsgBox "Please enter the room number", vbCritical, "Connex Event Management"
    data1.Recordset.cancel
    drive = 1
ElseIf date1.Value = "" Then
  MsgBox "Please enter the date the event", vbCritical, "Connex Event Management"
    data1.Recordset.cancel
    drive = 1
ElseIf date2.Value = "" Then
  MsgBox "Please enter the date the event", vbCritical, "Connex Event Management"
    data1.Recordset.cancel
    drive = 1
ElseIf time_in1.Text = "" Then
  MsgBox "Please enter the time that the event starts", vbCritical, "Connex Event Management"
    data1.Recordset.cancel
    drive = 1

ElseIf hour_type1.Text = "" Then
  MsgBox "Please enter the time period that the event is occuring", vbCritical, "Connex Event Management"
    data1.Recordset.cancel
    drive = 1
ElseIf group1.Text = "" Then
  MsgBox "Please enter the group hosting the event", vbCritical, "Connex Event Management"
    data1.Recordset.cancel
    drive = 1
ElseIf time_out1.Text = "" Then
  MsgBox "Please enter the time the event ends", vbCritical, "Connex Event Management"
    data1.Recordset.cancel
    drive = 1
ElseIf event_admin1.Text = "" Then
  MsgBox "Please enter Event Administrator's Name", vbCritical, "Connex Event Management"
    data1.Recordset.cancel
    drive = 1




End If

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  
  On Error GoTo has_error
  'This will display the current record position for this recordset
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
  
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


Private Sub cmdClose_Click()

On Error GoTo has_error
  Unload Me
  
  Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub










