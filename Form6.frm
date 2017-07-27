VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Entry Form"
   ClientHeight    =   4650
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7290
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   960
      OleObjectBlob   =   "Form6.frx":0000
      TabIndex        =   10
      Top             =   2400
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "Form6.frx":0066
      Top             =   1200
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   120
      Picture         =   "Form6.frx":2A64F
      ScaleHeight     =   435
      ScaleWidth      =   6675
      TabIndex        =   9
      Top             =   120
      Width           =   6735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Room Administrator"
      Height          =   1695
      Left            =   2760
      TabIndex        =   8
      Top             =   1560
      Width           =   4455
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "Form6.frx":2D4F4
         TabIndex        =   13
         Top             =   1200
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   360
         OleObjectBlob   =   "Form6.frx":2D55C
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "Form6.frx":2D5DC
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox phone 
         DataField       =   "Admin_Phone"
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox email 
         DataField       =   "Admin_Email"
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox admin_name 
         DataField       =   "Admin_Name"
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Reset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Edit Panel"
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   3600
      Width           =   7095
   End
   Begin VB.TextBox room 
      DataField       =   "Room_#"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Menu file_1 
      Caption         =   "File"
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu record_status 
      Caption         =   "Record Status"
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
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Dim record_control2 As Integer
Dim opps As Integer




Private Sub Add_Click()

On Error GoTo has_error

'************************************************************************************8
'***********************************************************************************
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
.Source = "Select * From Event_Room Order by [Room_#]"
End With

'***********************************************************************
'************************************************************************

Dim drive As Integer
drive = 0
check_form drive
If drive = 0 Then

If (record_control2 = 0) Then

With data1
.Open
.AddNew
.Fields("Room_#") = room.Text
.Fields("Admin_Name") = admin_name.Text
.Fields("Admin_Phone") = phone.Text
.Fields("Admin_Email") = email.Text
.Update
.Close
End With
record_control2 = -1
Add.Enabled = False
Reset.Enabled = True
room.Enabled = False
room.Enabled = False
admin_name.Enabled = False
phone.Enabled = False
email.Enabled = False
opps = -1
End If

If record_control2 = 1 Then

data1.Open
data1.Find "[Room_#]= '" & room.Text & "'"
data1.Delete
data1.AddNew
data1.Fields("Room_#") = room.Text
data1.Fields("Admin_Name") = admin_name.Text
data1.Fields("Admin_Phone") = phone.Text
data1.Fields("Admin_Email") = email.Text
data1.Update
data1.Close


record_control2 = -1
Add.Enabled = False
Reset.Enabled = False
room.Enabled = False
room.Enabled = False
admin_name.Enabled = False
phone.Enabled = False
email.Enabled = False
new_record.Enabled = True
edit_record.Enabled = False
search_record.Enabled = True
opps = -1
End If
End If
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub Reset_Click()

On Error GoTo has_error
record_control2 = -1
room.Enabled = False
admin_name.Enabled = False
phone.Enabled = False
email.Enabled = False
Add.Enabled = False
Delete.Enabled = False
Reset.Enabled = False
edit_record.Enabled = False
new_record.Enabled = True
search_record.Enabled = True
opps = -1
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub





Private Sub Delete_Click()

On Error GoTo has_error
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
.Source = "Select * From Event_Room Order by [Room_#]"
End With
'*********************************
With data1
.Open
.Find "[Room_#]= '" & room.Text & "'"
.Delete
.Close
End With
record_control2 = -1
Add.Enabled = False
Reset.Enabled = False
room.Enabled = False
room.Enabled = False
admin_name.Enabled = False
phone.Enabled = False
email.Enabled = False
new_record.Enabled = True
edit_record.Enabled = False
search_record.Enabled = True
Delete.Enabled = False
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub edit_record_Click()

On Error GoTo has_error
new_record.Enabled = False
search_record.Enabled = False
edit_record.Enabled = False

record_control2 = 1
Add.Enabled = True
Delete.Enabled = False
room.Enabled = True
admin_name.Enabled = True
phone.Enabled = True
email.Enabled = True
opps = 1
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
.Source = "Select * From Event_Room Order by [Room_#]"
End With


room.Enabled = False
admin_name.Enabled = False
phone.Enabled = False
email.Enabled = False
Add.Enabled = False
Delete.Enabled = False
Reset.Enabled = False
edit_record.Enabled = False
opps = -1
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub new_record_Click()

On Error GoTo has_error
new_record.Enabled = False
search_record.Enabled = False
edit_record = False
record_control2 = 0
room.Enabled = True
admin_name.Enabled = True
phone.Enabled = True
email.Enabled = True
Add.Enabled = True
Reset.Enabled = True
opps = 1
room.Text = ""
admin_name.Text = ""
phone.Text = ""
email.Text = ""
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub clear_box()

On Error GoTo has_error

admin_name.Text = ""
phone.Text = ""
email.Text = ""

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub search_record_Click()

On Error GoTo has_error
search_e.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub form_queryUnload(cancel As Integer, unloadmode As Integer)
On Error GoTo has_error
'************************************************************************************8
'***********************************************************************************
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
.Source = "Select * From Event_Room Order by [Room_#]"
End With

'***********************************************************************
'************************************************************************
With data1
.Open
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
.Close
End With

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub check_form(ByRef drive As Integer)
On Error GoTo has_error


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
.Source = "Select * From Event_Room Order by [Room_#]"
End With

'***********************************************************************
'************************************************************************

drive = 0
If room.Text = "" Then
    MsgBox "Please enter room name or number", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf admin_name.Text = "" Then
  MsgBox "Please enter room administrator", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf phone.Text = "" Then
  MsgBox "Please enter phone number", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf email.Text = "" Then
  MsgBox "Please enter email", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
End If

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

