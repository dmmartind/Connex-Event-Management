VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   570
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2040
      OleObjectBlob   =   "search_e6.frx":0000
      Top             =   0
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////




Option Explicit


Private Sub Form_Activate()

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
.Source = "Select * From Schedule_Info Order by [Event] ASC"
.Open
End With

With data1


Do Until data1.EOF
    If ![Event] <> "" Then
        Combo1.AddItem ![Event]
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
'********************************************
On Error GoTo has_error
Dim connection_string As String
connection_string = Form1.connection_string
Dim temp1 As String
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
.Open
End With
'********************************************

With data1
.Find "[Event]= '" & Combo1.Text & "'"
If .EOF = False Then
Form20.phone1 = ![phone#]
Form20.event1.Text = ![Event]
Form20.room1.Text = ![room]
Form20.date1.Value = ![Begin Date]
temp1 = ![Begin Date]
Form20.date2.Value = ![End Date]
Form20.time_in1.Text = ![Time]
Form20.time_out1.Text = ![time_out]
Form20.group1.Text = ![Name/Group]
Form20.hour_type1.Text = ![hour_type]
Form20.event_admin1.Text = ![administrator]
Form20.Time_Date.Text = ![Date_Time]
Form20.hour_type1.Enabled = False
Form20.group1.Enabled = False
Form20.time_out1.Enabled = False
Form20.event_admin1.Enabled = False
Form20.search_record.Enabled = True
Else
MsgBox "No record found!!!", vbExclamation, "Connex Event Management"
End If
.Close
End With
adn.Close

'///////////////////////////////////
'//////////////////////////////////
'Form2.DataGrid1.Enabled = True
    

Dim adn1 As ADODB.Connection
Dim data2 As ADODB.Recordset

Set adn1 = New ADODB.Connection
Set data2 = New ADODB.Recordset

With adn1
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data2
.ActiveConnection = adn1
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From activity2 where [Event]= '" & Combo1.Text & "' AND [Event_Date]= #" & temp1 & "#"
.Open
End With
'/////////////////////////////////////
'/////////////////////////////////////
Form20.DataGrid1.ClearFields
Form20.DataGrid1.Enabled = True
'data2.Find "[Event_Date]= '" & Form20.date1.Value & "'"
If data2.EOF = False Then
Set Form20.DataGrid1.DataSource = data2
Form20.DataGrid1.ReBind
'Form20.DataGrid1.Enabled = True



Else
MsgBox "No record found!!!", vbExclamation, "Connex Event Management"
End If


Unload Me

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me


End Sub

