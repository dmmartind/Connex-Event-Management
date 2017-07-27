VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Search_e4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2040
      OleObjectBlob   =   "Search_e4.frx":0000
      Top             =   240
   End
   Begin VB.CommandButton OK 
      Caption         =   "Search"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Search_e4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Option Explicit
Dim connection_string As String
Dim WithEvents adn20 As ADODB.Connection
Attribute adn20.VB_VarHelpID = -1
Dim WithEvents data2 As ADODB.Recordset
Attribute data2.VB_VarHelpID = -1







Private Sub Form_Activate()

On Error GoTo has_error
Skin1.ApplySkin Me.hwnd

Dim connection_string As String

connection_string = Form1.connection_string

'//////////////////////////////////////////////////////////////////////////////

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
.Source = "Select * From Schedule_Info Order by [Event]"
End With

With data1
.Open

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

On Error GoTo has_error
'//////////////////////////////////////////////////////////////////////////////
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
.Source = "Select * From Schedule_Info Order by [Event]"
End With



With data1
.Open
.Find "[Event]= '" & Combo1.Text & "'"
If .EOF = False Then
Form2.event1.Text = ![Event]
Form2.room1.Text = ![room]
Form2.date1.Value = ![Begin Date]
Form2.date2.Value = ![End Date]
Form2.time_in1.Text = ![Time]
Form2.time_out1.Text = ![time_out]
Form2.group1.Text = ![Name/Group]
Form2.hour_type1.Text = ![hour_type]
Form2.event_admin1 = ![administrator]
Form2.datetime = ![Date_Time]


Form2.date1.Enabled = False
Form2.date2.Enabled = False

Form2.hour_type1.Enabled = False
Form2.group1.Enabled = False
Form2.time_out1.Enabled = False
Form2.event_admin1.Enabled = False
Form2.new_record.Enabled = False
Form2.search_record.Enabled = False
Form2.edit_record = True
Form2.Add.Enabled = False
Form2.Delete.Enabled = True
Form2.Reset.Enabled = True
Form2.cmdAdd = False
Form2.cmdDelete = False
Form2.cmdnext = False
Form2.cmdRefresh = False
Form2.cmdUpdate = False

Else
MsgBox "No record found!!!", vbExclamation, "Connex Event Management"
End If
.Close
End With
adn.Close

Unload Me
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub


