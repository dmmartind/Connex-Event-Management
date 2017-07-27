VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Dialog2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activity Search"
   ClientHeight    =   675
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   840
      OleObjectBlob   =   "Form24.frx":0000
      Top             =   120
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Option Explicit



Private Sub OKButton_Click()
On Error GoTo has_error
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

With data1
.Open
If Combo1.Text <> "" Then
Dialog1.Text1.Text = Combo1.Text
Dialog1.hold = Combo1.Text
Dialog1.Delete.Enabled = True
Dialog1.Reset.Enabled = True
Dialog1.OKButton.Enabled = False
Dialog1.edit_record.Enabled = True
Dialog1.new_record.Enabled = False
Dialog1.search_record.Enabled = False
Else
MsgBox "You didn't search for anything!!!", vbExclamation, "Connex Event Management"
Dialog1.Delete.Enabled = False
Dialog1.Reset.Enabled = False
Dialog1.OKButton.Enabled = False
Dialog1.edit_record.Enabled = False
Dialog1.new_record.Enabled = True
Dialog1.search_record.Enabled = True
End If
.Close
End With

adn1.Close
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

With data1
.Open

Do Until data1.EOF
    If ![activity] <> "" Then
        Combo1.AddItem ![activity]
    End If
        .MoveNext
Loop
.Close
End With

adn1.Close

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub


