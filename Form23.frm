VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Dialog1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activity Populator"
   ClientHeight    =   1260
   ClientLeft      =   2760
   ClientTop       =   4035
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Reset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2760
      OleObjectBlob   =   "Form23.frx":0000
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "Form23.frx":2A5E9
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Save"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Menu record_control 
      Caption         =   "Record Control"
      Begin VB.Menu new_record 
         Caption         =   "New Record"
      End
      Begin VB.Menu edit_record 
         Caption         =   "Edit Record"
      End
      Begin VB.Menu search_record 
         Caption         =   "Search Record"
      End
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Option Explicit
Dim edit_form As Boolean
Public hold As String

Private Sub Delete_Click()

Delete.Enabled = False
edit_record.Enabled = False


'//////////////////////////////////////////////////
On Error GoTo h_run
'/////////////////////////////////////////////////


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
.Source = "Select * From activity_type where [Activity]= '" & Text1.Text & "'"
End With

With data1
.Open
.Find "[Activity]= '" & Text1.Text & "'"
If .EOF = False Then
.Delete
End If
.Close
End With
adn1.Close
Exit Sub

h_run:
MsgBox "Your delete request was denied", vbExclamation, "Connex Event Management"
End Sub

Private Sub edit_record_Click()
On Error GoTo has_error

OKButton.Enabled = True
Delete.Enabled = False
edit_record.Enabled = False
Text1.Enabled = True
edit_form = True
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub
Private Sub Form_Load()
On Error GoTo has_error
Skin1.ApplySkin Me.hwnd
Text1.Enabled = False
edit_record.Enabled = False
OKButton.Enabled = False
Delete.Enabled = False
Reset.Enabled = False
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me




End Sub

Private Sub new_record_Click()
On Error GoTo has_error
OKButton.Enabled = True
Reset.Enabled = True
Delete.Enabled = False

new_record.Enabled = False
search_record.Enabled = False
edit_record.Enabled = False
Text1.Enabled = True

Text1.Text = ""
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub OKButton_Click()
On Error GoTo has_error
OKButton.Enabled = False

'//////////////////////////////////////////////////////////
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
.Source = "Select * From activity_type"
End With

If Text1.Text <> "" Then

If edit_form = False Then



With data1
.Open
.AddNew
.Fields(1) = Text1.Text
.Update
.Close
End With
Text1.Enabled = False

Else

With data1
.Open
.Find "[Activity]= '" & hold & "'"
If .EOF = False Then
.Delete
End If
.Close
End With

With data1
.Open
.AddNew
.Fields(1) = Text1.Text
.Update
.Close
End With
Text1.Enabled = False
End If
Else
MsgBox "You need to type an activity!!!", vbExclamation, "Connex Event Management"
End If

adn1.Close
edit_form = False

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub Reset_Click()

On Error GoTo has_error
OKButton.Enabled = False
Delete.Enabled = False
Reset.Enabled = False





new_record.Enabled = True
search_record.Enabled = True
edit_record.Enabled = False

Text1.Enabled = False

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub search_record_Click()

On Error GoTo has_error
Dialog2.Show vbModal

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me


End Sub
