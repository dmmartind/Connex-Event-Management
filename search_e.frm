VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form search_e 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Employee"
   ClientHeight    =   585
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2040
      OleObjectBlob   =   "search_e.frx":0000
      Top             =   0
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "search_e"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Option Explicit
Dim connection_string As String






Private Sub Form_Activate()

On Error GoTo has_error
Skin1.ApplySkin Me.hwnd

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

Do Until data1.EOF
    If ![Room_#] <> "" Then
        Combo1.AddItem ![Room_#]
    End If
        .MoveNext
Loop
.Close
End With
adn.Close

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub OK_Click()

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
.Source = "Select * From Event_Room Order by [Room_#] ASC"
End With

'***********************************************************************
'************************************************************************

With data1
.Open
.Find "[Room_#]= '" & Combo1.Text & "'"
If .EOF = False Then
Form6.room.Text = ![Room_#]
Form6.admin_name.Text = ![admin_name]
Form6.phone.Text = ![Admin_Phone]
Form6.email.Text = ![Admin_Email]
Form6.room.Enabled = False
Form6.admin_name.Enabled = False
Form6.phone.Enabled = False
Form6.email.Enabled = False
Form6.edit_record.Enabled = True
Form6.new_record.Enabled = False
Form6.search_record.Enabled = False
Form6.edit_record.Enabled = True
Form6.new_record.Enabled = False
Form6.Add.Enabled = False
Form6.Delete.Enabled = True
Form6.Reset.Enabled = True
Else
MsgBox "No record found!!!", vbExclamation, "Connex Event Management"
End If
Unload Me
End With

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub
