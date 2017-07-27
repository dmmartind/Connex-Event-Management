VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Search_e1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2160
      OleObjectBlob   =   "Search_e1.frx":0000
      Top             =   120
   End
   Begin VB.CommandButton OK 
      Caption         =   "Search"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Search_e1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim connection_string As String






Private Sub Form_Activate()

On Error GoTo has_error
Skin1.ApplySkin Me.hwnd
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













'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


With data1
.Open


Do Until data1.EOF
    If ![Employee_Name] <> "" Then
        Combo1.AddItem ![Employee_Name]
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


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim connection_string As String

connection_string = Form1.connection_string
Dim adn As ADODB.Connection
Dim data1 As ADODB.Recordset
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

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

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


With data1
.Open
.Find "[Employee_Name]= '" & Combo1.Text & "'"
If .EOF = False Then
Form14.name1.Text = ![Employee_Name]
Form14.original_name = ![Employee_Name]
Form14.work_status1.Text = ![Work_Status]
Form14.phone1.Text = ![Phone_Number]
Form14.e_mail1.Text = ![email]
Form14.name1.Enabled = False
Form14.work_status1.Enabled = False
Form14.phone1.Enabled = False
Form14.e_mail1.Enabled = False
Form14.edit_record.Enabled = True
Form14.new_record.Enabled = False
Form14.search_record.Enabled = False
Form14.Add.Enabled = False
Form14.Delete.Enabled = True
Form14.Reset.Enabled = True
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

