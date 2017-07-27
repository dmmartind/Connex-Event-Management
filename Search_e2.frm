VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Search_e2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   4065
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1800
      OleObjectBlob   =   "Search_e2.frx":0000
      Top             =   120
   End
   Begin VB.CommandButton OK 
      Caption         =   "Search"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Search_e2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
.Source = "Select * From Customers Order by [Name]"
End With

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////



With data1
.Open

Do Until data1.EOF
    If ![name] <> "" Then
        Combo1.AddItem ![name]
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

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////

With data1
.Open
.Find "[Name]= '" & Combo1.Text & "'"
If .EOF = False Then
Form9.name1.Text = ![name]
Form9.phone1.Text = ![phone]
Form9.group1.Text = ![Group]
Form9.email1.Text = ![email]
Form9.name1.Enabled = False
Form9.phone1.Enabled = False
Form9.group1.Enabled = False
Form9.email1.Enabled = False
Form9.edit_record.Enabled = True
Form9.new_record.Enabled = False
Form9.search_record.Enabled = False
Form9.Add.Enabled = False
Form9.Delete.Enabled = True
Form9.Reset.Enabled = True
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

