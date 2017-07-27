VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Form14 
   Caption         =   "Employee Entry Form"
   ClientHeight    =   4650
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   495
      Left            =   3840
      OleObjectBlob   =   "Form5.frx":0000
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   3840
      OleObjectBlob   =   "Form5.frx":0080
      TabIndex        =   11
      Top             =   2280
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form5.frx":00EC
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "Form5.frx":015E
      TabIndex        =   9
      Top             =   2280
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5760
      OleObjectBlob   =   "Form5.frx":01C4
      Top             =   840
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   120
      Picture         =   "Form5.frx":2A7AD
      ScaleHeight     =   435
      ScaleWidth      =   6675
      TabIndex        =   8
      Top             =   120
      Width           =   6735
   End
   Begin VB.ComboBox work_status1 
      DataField       =   "Work_Status"
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3120
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
   Begin VB.TextBox e_mail1 
      DataField       =   "Email"
      Height          =   285
      Left            =   4800
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox phone1 
      DataField       =   "Phone_Number"
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox name1 
      DataField       =   "Employee_Name"
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Menu file_1 
      Caption         =   "File"
      Begin VB.Menu exit 
         Caption         =   "E&xit"
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
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Dim record_control2 As Integer
Dim opps As Integer
Public original_name As String




Private Sub Add_Click()

On Error GoTo has_error
Dim connection_string As String
'********************************************************************\
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
.Source = "Select * From Employee Order by [Employee_Name]"
End With

'*********************************************************************


Dim drive As Integer

check_form drive
If drive = 0 Then

If (record_control2 = 0) Then
With data1
.Open
.AddNew
.Fields("Employee_Name") = name1.Text
.Fields("Work_Status") = work_status1.Text
.Fields("Phone_Number") = phone1.Text
.Fields("Email") = e_mail1.Text
.Update
.Close
End With
add_in
End If


If record_control2 = 1 Then
add_in2
End If
End If

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub
Private Sub add_in()

On Error GoTo has_error

record_control2 = -1
Add.Enabled = False
Reset.Enabled = True
name1.Enabled = False
work_status1.Enabled = False
phone1.Enabled = False
e_mail1.Enabled = False
opps = -1

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me


End Sub
Private Sub add_in2()

On Error GoTo has_error
'********************************************************************\
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
.Source = "Select * From Employee Order by [Employee_Name]"
End With

'*********************************************************************
With data1
.Open
.Find "[Employee_Name]= '" & original_name & "'"
If .EOF = False Then
.Delete
.AddNew
.Fields("Employee_Name") = name1.Text
.Fields("Work_Status") = work_status1.Text
.Fields("Phone_Number") = phone1.Text
.Fields("Email") = e_mail1.Text
.Update
.Close
record_control2 = -1
Add.Enabled = False
Reset.Enabled = True
name1.Enabled = False
work_status1.Enabled = False
phone1.Enabled = False
e_mail1.Enabled = False
new_record.Enabled = True
edit_record.Enabled = False
search_record.Enabled = True
opps = -1
End If
End With

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub Reset_Click()

On Error GoTo has_error
record_control2 = -1
name1.Enabled = False
work_status1.Enabled = False
phone1.Enabled = False
e_mail1.Enabled = False
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
'**************************************
'***************************************
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
.Source = "Select * From Employee Where [Employee_Name]= '" & name1.Text & "' Order by [Employee_Name]"
End With
'*************************************
'*************************************

With data1
.Open
.Delete
.Close
End With
record_control2 = -1
Add.Enabled = False
name1.Enabled = False
work_status1.Enabled = False
phone1.Enabled = False
e_mail1.Enabled = False
new_record.Enabled = True
edit_record.Enabled = False
search_record.Enabled = True
Delete.Enabled = False
adn.Close

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub edit_record_Click()

On Error GoTo has_error
record_control2 = 1
edit_record.Enabled = False
Add.Enabled = True
Delete.Enabled = False
name1.Enabled = True
work_status1.Enabled = True
phone1.Enabled = True
e_mail1.Enabled = True
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
Dim adn1 As ADODB.Connection
Dim data2 As ADODB.Recordset
Dim data1 As ADODB.Recordset


Set adn = New ADODB.Connection
Set adn1 = New ADODB.Connection
Set data2 = New ADODB.Recordset
Set data1 = New ADODB.Recordset

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
.Source = "Select * From Employee Order by [Employee_Name]"
End With


With data2
.ActiveConnection = adn1
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Pay_Status"
End With

record_control2 = -1
name1.Enabled = False
work_status1.Enabled = False
phone1.Enabled = False
e_mail1.Enabled = False
Add.Enabled = False
Delete.Enabled = False
Reset.Enabled = False
edit_record.Enabled = False
opps = -1

With data2
.Open


Do Until data2.EOF
    If ![Pay_Status] <> "" Then
        work_status1.AddItem ![Pay_Status]
    End If
        .MoveNext
Loop
.Close
End With
adn.Close
adn1.Close
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub new_record_Click()

On Error GoTo has_error
new_record.Enabled = False
search_record.Enabled = False
record_control2 = 0
name1.Enabled = True
work_status1.Enabled = True
phone1.Enabled = True
e_mail1.Enabled = True
Add.Enabled = True
Reset.Enabled = True
opps = 1
name1.Text = ""
work_status1.Text = ""
phone1.Text = ""
e_mail1.Text = ""
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me


End Sub



Private Sub search_record_Click()

On Error GoTo has_error
Search_e1.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub form_queryUnload(cancel As Integer, unloadmode As Integer)
On Error GoTo has_error
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
.Source = "Select * From Employee Order by [Employee_Name]"
End With

'*********************************************************************




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
Unload Me
End If
.Close

End With
adn.Close

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
.Source = "Select * From Employee Order by [Employee_Name]"
End With











drive = 0
If name1.Text = "" Then
    MsgBox "Please enter name", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf work_status1.Text = "" Then
  MsgBox "Please enter pay status", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf phone1.Text = "" Then
  MsgBox "Please enter phone number", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
ElseIf e_mail1.Text = "" Then
  MsgBox "Please enter email", vbCritical, "Connex Event Management"
    data1.cancel
    drive = 1
End If

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me



End Sub




