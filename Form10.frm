VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User and Group Accounts"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   5640
   StartUpPosition =   1  'CenterOwner
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2640
      OleObjectBlob   =   "Form10.frx":0000
      Top             =   2520
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "User Permissions"
      TabPicture(0)   =   "Form10.frx":2A5E9
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Combo1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "submit"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "password"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "result"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Password"
      TabPicture(1)   =   "Form10.frx":2A605
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(4)=   "name1"
      Tab(1).Control(5)=   "old_password"
      Tab(1).Control(6)=   "New_Password"
      Tab(1).Control(7)=   "verify"
      Tab(1).Control(8)=   "submit2"
      Tab(1).ControlCount=   9
      Begin VB.TextBox result 
         DataField       =   "Group"
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton password 
         Caption         =   "Clear Password"
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton submit2 
         Caption         =   "Submit"
         Height          =   375
         Left            =   -72960
         TabIndex        =   9
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox verify 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -72720
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox New_Password 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -72720
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox old_password 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -72720
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox name1 
         Height          =   285
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   2655
      End
      Begin VB.Frame Frame1 
         Caption         =   "Permissions"
         Height          =   855
         Left            =   720
         TabIndex        =   13
         Top             =   2400
         Width           =   4095
         Begin VB.OptionButton user1 
            Caption         =   "User"
            Height          =   495
            Left            =   2880
            TabIndex        =   3
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton admin1 
            Caption         =   "Admin"
            Height          =   375
            Left            =   360
            TabIndex        =   2
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.CommandButton submit 
         Caption         =   "Submit"
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   3600
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Verify New Password"
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "New Password:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Old Password:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "User Name"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton ok 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   5040
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Option Explicit
Dim index As Integer
Dim connection_string As String

Private Sub Combo1_Click()
'**********************************************
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

'**********************************************







With data1
.Open
.Find "[Employee_Name] = '" & Combo1.Text & "'"
Select Case ![Group]
Case "Admin"
admin1.Value = True
user1.Value = False
Case "User"
admin1.Value = False
user1.Value = True
End Select
name1.Text = ![Employee_Name]
.Close
End With
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
.Source = "Select * From Employee Order by [Employee_Name]"
End With

With data1
.Open

Do Until .EOF
    If ![Employee_Name] <> "" Then
        Combo1.AddItem ![Employee_Name]
    End If
        .MoveNext
Loop
.Close
End With

name1.Text = Combo1.Text
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub OK_Click()
On Error GoTo has_error
Unload Me
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub password_Click()

On Error GoTo has_error

Dim adocn As ADODB.Connection
Set adocn = New ADODB.Connection
Dim connection_string As String

connection_string = Form1.connection_string


With adocn
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With


Dim data1 As ADODB.Recordset

Set data1 = New ADODB.Recordset

With data1
.ActiveConnection = adocn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Employee Order by [Employee_Name]"
End With


With data1
data1.Open
data1.MoveFirst
data1.Find "[Employee_Name] = '" & Combo1.Text & "'"
data1.Fields(6) = ""
data1.Update
data1.Close
End With

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub submit_Click()

On Error GoTo has_error
Dim connection_string As String
connection_string = Form1.connection_string

Dim adocn As ADODB.Connection
Set adocn = New ADODB.Connection

With adocn
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With


Dim data1 As ADODB.Recordset

Set data1 = New ADODB.Recordset

With data1
.ActiveConnection = adocn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Employee Order by [Employee_Name]"

End With

Rem ////////////////////////////////
With data1
data1.Open
data1.MoveFirst
.Find "[Employee_Name] = '" & Combo1.Text & "'"
Select Case admin1.Value
Case True
data1.Fields(5) = "Admin"
data1.Update
Case False
data1.Fields(5) = "User"
data1.Update
End Select
.Close
End With

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub submit2_Click()

On Error GoTo has_error
Dim connection_string As String
connection_string = Form1.connection_string

Dim adocn As ADODB.Connection
Set adocn = New ADODB.Connection

With adocn
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With


Dim data1 As ADODB.Recordset

Set data1 = New ADODB.Recordset

With data1
.ActiveConnection = adocn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Employee Order by [Employee_Name]"
End With







With data1
.Open
.Find "[Employee_Name] = '" & Combo1.Text & "'"
If old_password.Text = data1.Fields(6).Value Then
If New_Password.Text = verify.Text Then
data1.Fields(6) = verify.Text
.Update
Else
MsgBox "Please enter the correct new password in the verification box and prey submit", vbInformation, "Connex Event Management"
End If
Else
MsgBox "Please enter the correct old password in the old password box.", vbInformation, "Connex Event Management"
End If
.Close
End With

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub
