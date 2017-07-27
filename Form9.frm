VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Entry Form"
   ClientHeight    =   4650
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7305
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7305
   StartUpPosition =   1  'CenterOwner
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   375
      Left            =   3840
      OleObjectBlob   =   "Form9.frx":0000
      TabIndex        =   12
      Top             =   2880
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   3840
      OleObjectBlob   =   "Form9.frx":0080
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form9.frx":00F4
      TabIndex        =   10
      Top             =   2880
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form9.frx":0160
      TabIndex        =   9
      Top             =   2280
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3720
      OleObjectBlob   =   "Form9.frx":01C6
      Top             =   960
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   120
      Picture         =   "Form9.frx":2A7AF
      ScaleHeight     =   435
      ScaleWidth      =   6675
      TabIndex        =   8
      Top             =   120
      Width           =   6735
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Reset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   495
      Left            =   240
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
   Begin VB.TextBox email1 
      DataField       =   "EMail"
      Height          =   285
      Left            =   4920
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox group1 
      DataField       =   "Group"
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox phone1 
      DataField       =   "Phone"
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox name1 
      DataField       =   "Name"
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Menu file_1 
      Caption         =   "File"
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu record_status 
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
Attribute VB_Name = "Form9"
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
.Source = "Select * From Customers Order by [Name]"
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
.Fields("Name") = name1.Text
.Fields("Group") = group1.Text
.Fields("Phone") = phone1.Text
.Fields("EMail") = email1.Text
.Update
.Close
End With
record_control2 = -1
Add.Enabled = False
Reset.Enabled = True
name1.Enabled = False
phone1.Enabled = False
group1.Enabled = False
email1.Enabled = False
opps = -1
End If

If record_control2 = 1 Then
With data1
.Open
.Find "[Name]= '" & name1.Text & "'"
.Delete
.AddNew
.Fields("Name") = name1.Text
.Fields("Group") = group1.Text
.Fields("Phone") = phone1.Text
.Fields("EMail") = email1.Text
.Update
.Close
End With
record_control2 = -1
Add.Enabled = False
Reset.Enabled = True
name1.Enabled = False
phone1.Enabled = False
group1.Enabled = False
email1.Enabled = False
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
name1.Enabled = False
phone1.Enabled = False
group1.Enabled = False
email1.Enabled = False
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
.Source = "Select * From Customers Order by [Name] ASC"
End With

'***********************************************************************
'************************************************************************
With data1
.Open
.Find "[name]= '" & name1.Text & "'"
.Delete
End With
record_control2 = -1
Add.Enabled = False
Reset.Enabled = False
name1.Enabled = False
phone1.Enabled = False
group1.Enabled = False
email1.Enabled = False
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
edit_record.Enabled = False
record_control2 = 1
Add.Enabled = True
Delete.Enabled = False
name1.Enabled = True
phone1.Enabled = True
group1.Enabled = True
email1.Enabled = True
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
record_control2 = -1

Rem With data1
'.ConnectionString = connection_string
'.CommandType = adCmdTable
'.RecordSource = "Customers"

'End With


name1.Enabled = False
phone1.Enabled = False
group1.Enabled = False
email1.Enabled = False
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
record_control2 = 0

new_record.Enabled = False
search_record = False


name1.Enabled = True
phone1.Enabled = True
group1.Enabled = True
email1.Enabled = True
Add.Enabled = True
Reset.Enabled = True
opps = 1
name1.Text = ""
phone1.Text = ""
group1.Text = ""
email1.Text = ""
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub


Private Sub search_record_Click()

On Error GoTo has_error
Search_e2.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub form_queryUnload(cancel As Integer, unloadmode As Integer)
'************************************************************************************8
'***********************************************************************************


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
.Source = "Select * From Customers Order by [Name]"
End With

'***********************************************************************
'************************************************************************







Dim temp As Integer
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
.Source = "Select * From Customers Order by [Name]"
End With

drive = 0
If name1.Text = "" Then
    MsgBox "Please enter name", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf phone1.Text = "" Then
  MsgBox "Please enter phone number", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf group1.Text = "" Then
  MsgBox "Please enter the group or department name.", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf email1.Text = "" Then
  MsgBox "Please enter email or Alturnative phone number.", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
End If

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

