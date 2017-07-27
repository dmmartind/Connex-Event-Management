VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Log-In"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1680
      OleObjectBlob   =   "frmLogin.frx":0000
      Top             =   480
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmLogin.frx":2A5E9
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmLogin.frx":2A657
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   0
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   2
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   3
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Option Explicit

Public LoginSucceeded As Boolean
Public connection_string As String


Private Sub cmdCancel_Click()
On Error GoTo has_error
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
    
   Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"

 Unload Me
End Sub

Private Sub cmdOK_Click()

On Error GoTo has_error
    'check for correct password
   Dim connection_string As String
    connection_string = Form1.connection_string

Dim adn As ADODB.Connection
Dim data1 As ADODB.Recordset

Set adn = New ADODB.Connection
Set data1 = New ADODB.Recordset

With adn
.connectionstring = connection_string
.CursorLocation = adUseClient
On Error GoTo Handle

.Open
End With


With data1
.ActiveConnection = adn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Employee Order by [Employee_Name]"
.Open
End With

If txtUserName.Text = "admin" Then
         If txtPassword.Text = "W4n$mOT*tb6L-)x)IqQ4" Then
        Form1.unlock_admin2
        Else
        MsgBox "Password is incorrect", vbExclamation, "Connex Event Management"
        End If
 Else
       With data1
        .Find "[employee_name]= '" & txtUserName.Text & "'"
        If .EOF = False Then
               
        If data1![Employee_Name] = txtUserName.Text Then
            If ![password] = txtPassword.Text Then
                If ![Group] = "Admin" Then
                    Form1.user_name1 = txtUserName.Text
                    Form1.unlock_admin
                ElseIf ![Group] = "User" Then
                    Form1.user_name1 = txtUserName.Text
                    Form1.unlock_user
                Else
                    MsgBox "You have not been authorized yet to use this program. Check software administrator.", vbExclamation, "Connex Event Management"
                End If
        Else
        MsgBox "Password is incorrect!!! Please try again!!!", vbExclamation, "Connex Event Management"
        
            End If
        
        Else
            MsgBox "Either your password did not match or you have not been authorized to use this software. Check with software administrator.", vbExclamation, "Connex Event Management"
        End If
        Else
        MsgBox "You are not found as a current user", vbCritical, "Connex Event Management"
        End If
        
        .Close
        
        End With
        
        End If
        
        
Unload Me
Exit Sub

Handle:
Dim ike As Integer
ike = MsgBox("This program requires that the database,that was supplied in the installation, be installed prior to running this program for the first time. If you have not installed the database, please exit now. If you have installed the database, please be sure to select the directory, that the database resides in, in the settings dialog. NOTE: If the database resides on a network you must map the drive prior to running this program.", vbCritical, "Connex Event Management")
Select Case ike
Case vbOK
Unload Me
Case vbCancel
Rem Dim OneForm As Form
Rem For Each OneForm In Forms
Rem Unload OneForm
Rem Next
Unload Me
End Select

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"

Unload Me





End Sub

Private Sub Form_Load()

Skin1.ApplySkin Me.hwnd

Rem connection_string = Form1.connection_string
Rem With data1
Rem .ConnectionString = connection_string
Rem .CommandType = adCmdTable
Rem .RecordSource = "Employee"
Rem .Refresh
Rem End With
End Sub
