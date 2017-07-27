VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form room_report 
   Caption         =   "Event Room Report"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdpreview 
      Caption         =   "Preview Report"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Date Location"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   6615
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "room_report.frx":0000
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "room_report.frx":006E
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox begin_time 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox end_time 
         Height          =   315
         Left            =   4800
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Room Location"
      Height          =   2775
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   240
         Width           =   3615
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5520
      OleObjectBlob   =   "room_report.frx":00E0
      Top             =   1200
   End
End
Attribute VB_Name = "room_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Private Sub cmdpreview_Click()

On Error GoTo has_error
Dim connection_string As String
connection_string = Form1.connection_string

Dim adn1 As ADODB.Connection
Dim adn2 As ADODB.Connection
Dim data1 As ADODB.Recordset
Dim data2 As ADODB.Recordset

Set adn1 = New ADODB.Connection
Set adn2 = New ADODB.Connection
Set data1 = New ADODB.Recordset
Set data2 = New ADODB.Recordset

With adn1
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With adn2
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With
With data1
.ActiveConnection = adn1
.CursorType = adOpenDynamic
.LockType = adLockOptimistic
.Source = "SELECT * FROM Schedule_Info Order by [Event] ASC"
End With

Dim reboot As Integer

Dim count As Integer
count = 0
'///////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////

Dim name As String
name = App.Path & "\test.html"
Open name For Output As #1

Print #1, "<HTML>"
Print #1, "<HEAD>"
Print #1, "<TITLE>Database Report</TITLE>"
Print #1, "</HEAD>"
Print #1, "<BODY bgcolor=""#f3e7b1"">"
Print #1, "<CENTER>"
Print #1, "<STRONG><FONT face=""times new roman"" color=""#000000"" size=""7"">Room Statistics Report</FONT></STRONG>"
Print #1, "</CENTER>"
Print #1, "<HR>"
Print #1, "<br>"
Print #1, " <br>"
Print #1, "<TABLE bgcolor=""#f3e7b1"">"
Print #1, "<TR>"
Print #1, "</TABLE>"
'/////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////
Dim flag As Boolean


reboot = 1
While count <> List1.ListCount
flag = False
List1.ListIndex = count
With data1
.Open
If List1.Selected(count) Then
reboot = 1
'//////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////
While (flag = False)
.Close
.Source = "SELECT * FROM Schedule_Info Where [Room]= '" & List1.Text & "'"
.Open
If .EOF = False Then
reboot = 0
Print #1, "<TABLE bgcolor=""#f3e7b1"">"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""5"">Room</font></strong></TD>"
Print #1, "<TR>"
Print #1, "<TD bgcolor=""#f3e7b1"">", List1.Text, " </TD>"
Print #1, "</TR>"
Print #1, "<TR>"
Print #1, "<TD></TD>"
Print #1, "<TD></TD>"
Print #1, "</TR>"
Print #1, "</TABLE>"
Print #1, "<center>"
Print #1, "<TABLE bgcolor=""#f3e7b1"">"
Print #1, "<TR>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Events</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Hours</font></strong></TD>"
Print #1, "</TR>"
Print #1, "</TABLE>"
Print #1, "</center>"
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
While (.EOF = False)
If .EOF = False Then
'////////////////////////////////////////////////////////////////////////////////////////////////////////


If ((![Begin Date] <= begin_time.Text) And (![End Date] >= end_time.Text)) Then
Print #1, "<center>"
Print #1, "<TABLE bgcolor=""#f3e7b1"">"
Print #1, "<TR>"
Print #1, "</TR>"
Print #1, "<TR>"
Print #1, "<TD></TD>"
Print #1, "<TD></TD>"
Print #1, "</TR>"
Print #1, "<TR>"
Print #1, "<TD>", ![Event], "</TD>"
Print #1, "<TD>", ![total_hours], "</TD>"
Print #1, "</TR>"
Print #1, "<TR>"
Print #1, "<TD></TD>"
Print #1, "<TD></TD>"
Print #1, "</TR>"
Print #1, "</TABLE>"
Print #1, "</center>"
End If
End If
.MoveNext
Wend
flag = True
Print #1, "</TR>"
Print #1, "<TR>"
Print #1, "<TD></TD>"
Print #1, "<TD></TD>"
Print #1, "</TR>"
Print #1, "</TABLE>"
Print #1, "</center>"
If .EOF = False Then
.MoveNext
End If
Else
flag = True
Print #1, "</TR>"
Print #1, "<TR>"
Print #1, "<TD></TD>"
Print #1, "<TD></TD>"
Print #1, "</TR>"
Print #1, "</TABLE>"
Print #1, "</center>"

End If

Wend

'//////////////////////////////////////////////////////////////////////////////////////////////////////////
If .EOF = True And reboot = 1 Then
Print #1, "<TABLE>"
Print #1, "<TR>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""5"">Room</font></strong></TD>"
Print #1, "<TR>"
Print #1, "<TD bgcolor=""#f3e7b1"">", List1.Text, " </TD>"
Print #1, "</TR>"
Print #1, "<TR>"
Print #1, "<TD></TD>"
Print #1, "<TD></TD>"
Print #1, "</TR>"
Print #1, "</TABLE>"
End If
'///////////////////////////////////////////////////////////////////////////////////////////////////////////


End If


count = count + 1
.Close
End With

Wend

'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////
Print #1, " </TABLE>"
Print #1, "<P>"
Print #1, "<BR>"
Print #1, "<center>&copy;1989-2003 UHD Multimedia Services Central Event Managment System/David Martin<center>"
Print #1, "<center>All Rights Reserved<center>"
Print #1, "</BODY></HTML>"
Close #1

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
Dim adn2 As ADODB.Connection
Dim data1 As ADODB.Recordset
Dim data2 As ADODB.Recordset

Set adn1 = New ADODB.Connection
Set adn2 = New ADODB.Connection
Set data1 = New ADODB.Recordset
Set data2 = New ADODB.Recordset

With adn1
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With adn2
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data1
.ActiveConnection = adn1
.CursorType = adOpenDynamic
.LockType = adLockOptimistic
.Source = "SELECT [Room_#] FROM Event_Room Order By [Room_#]"
End With


With data2
.ActiveConnection = adn2
.CursorType = adOpenDynamic
.LockType = adLockOptimistic
.Source = "SELECT * FROM Schedule_Info Order by [Event] ASC"
End With

With data1
.Open
.MoveFirst
While .EOF = False
List1.AddItem ![Room_#]
.MoveNext
Wend
.Close
End With

With data2
.Open
If .EOF = False Then
.Sort = "[Begin Date]"
End If
While .EOF = False
begin_time.AddItem ![Begin Date]
.MoveNext
Wend
.Close
.Open

If .EOF = False Then
.Sort = "[End Date]"
End If

While .EOF = False
end_time.AddItem ![End Date]
.MoveNext
Wend
.Close
End With
adn1.Close
adn2.Close
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub
