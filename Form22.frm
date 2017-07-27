VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Form5 
   Caption         =   "Schedule Report"
   ClientHeight    =   4920
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8640
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton preview1 
      Caption         =   "Preview"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   4440
      Width           =   2655
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7680
      OleObjectBlob   =   "Form22.frx":0000
      Top             =   1920
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "Form22.frx":2A5E9
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "Form22.frx":2A65F
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Statistics"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   8415
      Begin VB.CheckBox employee 
         Caption         =   "Employee List"
         Height          =   255
         Left            =   3120
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox ev_ent 
         Caption         =   "# of events based on time-status"
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox activity 
         Caption         =   "Activity"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox hour1 
         Caption         =   "Hours Total"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox details 
         Caption         =   "Details"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "Form22.frx":2A6D9
      Left            =   1800
      List            =   "Form22.frx":2A6DB
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   1560
      Width           =   5055
   End
   Begin VB.ComboBox end1 
      Height          =   315
      Left            =   6960
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox begin 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







Private Sub preview1_Click()
On Error GoTo has_error
Dim connection_string As String
connection_string = Form1.connection_string
Dim greatest As Boolean



Dim name As String
name = App.Path & "\test.html"
Open name For Output As #1



Dim data2_book As Variant
Dim data10_book As Variant

Dim adn As ADODB.Connection
Dim data1 As ADODB.Recordset
Dim adn1 As ADODB.Connection
Dim data2 As ADODB.Recordset
Dim adn2 As ADODB.Connection
Dim data9 As ADODB.Recordset


Set adn = New ADODB.Connection
Set data1 = New ADODB.Recordset
Set adn1 = New ADODB.Connection
Set data2 = New ADODB.Recordset
Set adn2 = New ADODB.Connection
Set data9 = New ADODB.Recordset

With adn
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

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
.ActiveConnection = adn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Schedule_Info Order By [Event] ASC"
End With

With data2
.ActiveConnection = adn1
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From activity2 Order by [Event] ASC"
End With

With data9
.ActiveConnection = adn2
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Event_Time"
End With





Rem Set DataGrid1.DataSource = data2
Rem DataGrid1.ReBind


'//////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////


Print #1, "<HTML>"
Print #1, "<HEAD>"
Print #1, "<TITLE>Database Report</TITLE>"
Print #1, "</HEAD>"
Print #1, "<BODY bgcolor=""#f3e7b1"">"
Print #1, "<CENTER>"
Print #1, "<STRONG><FONT face=""times new roman"" color=""#000000"" size=""7"">Schedule Report</FONT></STRONG>"
Print #1, "</CENTER>"
Print #1, "<HR>"
Print #1, "<br>"
Print #1, " <br>"
Dim count As Integer
Dim find_count As Integer
count = 0
While count < List1.ListCount
Print #1, "<TABLE bgcolor=""#f3e7b1"">"
Print #1, "<TR>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""5"">Event</font></strong></TD>"

If List1.Selected(count) = True Then
find_count = find_count + 1
Print #1, "<TR>"
Print #1, "<TD bgcolor=""#f3e7b1"">", List1.List(count), " </TD>"
Print #1, "</TR>"
Print #1, "<TR>"
Print #1, "<TD></TD>"
Print #1, "<TD></TD>"
Print #1, "</TR>"
Print #1, "</TABLE>"
Print #1, "<center>"
Print #1, "<TABLE bgcolor=""#f3e7b1"">"
End If


'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
If details.Value = 1 Then
With data1
.Open
.MoveFirst
.Find "[Event]= '" & List1.List(count) & "'"
If .EOF = False Then
Print #1, "<TR>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Time and Date Entered</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Starting Date</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Ending Date</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Begining Time</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Ending Time</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Time Status</font></strong></TD>"
repeat_html 1
Print #1, "<TR>"
Print #1, "<TD>", ![Date_Time], "</TD>"
Print #1, "<TD>", ![Begin Date], "</TD>"
Print #1, "<TD>", ![End Date], "</TD>"
Print #1, "<TD>", ![Time], "</TD>"
Print #1, "<TD>", ![time_out], "</TD>"
Print #1, "<TD>", ![hour_type], "</TD>"
Print #1, "</TR>"
repeat_html 1
Print #1, "<TR>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Event Admin</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Time Status</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Group Dept.</font></strong></TD>"
Print #1, "</TR>"
repeat_html 1
Print #1, "<TR>"
Print #1, "<TD>", ![administrator], "</TD>"
Print #1, "<TD>", ![hour_type], "</TD>"
Print #1, "<TD>", ![Name/Group], "</TD>"
Print #1, "</TR>"
repeat_html 1
End If
.Close
End With
End If
'///////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////


'/////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
If activity.Value = 1 Then
Dim count_m As Integer
Dim count_n As Integer

Dim x As Integer
x = 0

'////////////////////////////////////////////////////////////////////////

count_m = -1
With data2
'greatest = greater_than(![event_date], begin.Text)


'.Source = "Select * From activity2 Where [Event]= '" & List1.List(count) & "' BETWEEN '" & begin.Text & "' AND '" & end1.Text & "'"
.Source = "Select * From activity2 Where [Event]= '" & List1.List(count) & "' AND [Event_Date] >= #" & begin.Text & "#"

.Open
'.Filter = "[Event]= '" & List1.List(count) & "' AND [Event_Date] >= '" & begin.Text & "'"

'.Find "[Event]= '" & List1.List(count) & "'"
'.Find "[Event_Date] >= '" & begin.Text & "'"
If .EOF = False Then

Print #1, "<TR>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Activity</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Equipment</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Problems</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Employee</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Hours</font></strong></TD>"
Print #1, "</TR>"
While data2.EOF = False
count_m = count_m + 1
count_n = count_n + 1
Print #1, "<TR>"
Print #1, "<TD>", ![activity], "</TD>"
Print #1, "<TD>", ![Equipment], "</TD>"
Print #1, "<TD>", ![Problems], "</TD>"
'Print #1, "</TR>"
'repeat_html 1
'///////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////


Dim temp1 As String
Dim temp2 As String

temp1 = ![Event]
temp2 = ![event_date]


If employee.Value = 1 Then
funk1 temp1, temp2, count_m
End If

.MoveNext
Wend

'////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////
'Print #1, "</TR>"
'count = count + 1
End If
.Close
End With
End If
'/////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////

'///////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////
If hour1.Value = 1 Then
With data1
.Open
.MoveFirst
.Find "[Event]= '" & List1.List(count) & "'"
If .EOF = False Then
Print #1, "<TR>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Total Hours</font></strong></TD>"
Print #1, "</TR>"
repeat_html 1
Print #1, "<TR>"
Print #1, "<TD>", ![total_hours], "</TD>"
Print #1, "</TR>"
End If
.Close
End With
End If
repeat_html 30
Print #1, "<TR>"
count = count + 1
Wend

If ev_ent.Value = 1 Then
'/////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////
Dim Week_Day As Integer
Dim Weekend_Day As Integer
Dim Weekend_Night As Integer
Dim Week_Day_Night As Integer
Dim Weekend_Day_Night As Integer
Week_Day = 0
Week_Night = 0
Weekend_Day = 0
Weekend_Night = 0
Week_Day_Night = 0
Weekend_Day_Night = 0
'//////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////

count = 0
While count < List1.ListCount
If List1.Selected(count) = True Then
find_count = find_count + 1
data9.Open
data9.MoveFirst
While data9.EOF = False
data1.Source = "Select * From Schedule_Info where [Hour_Type]= '" & data9![Event_Time] & "'"
data1.Open
data1.Find "[Event]= '" & List1.List(count) & "'"
If data1.EOF = False Then
Select Case data9![Event_Time]
Case "Week Day"
Week_Day = Week_Day + 1
Case "Week Night"
Week_Night = Week_Night + 1
Case "Weekend Day"
Weekend_Day = Weekend_Day + 1
Case "Weekend Night"
Weekend_Night = Weekend_Night + 1
Case "Week Day-Night"
Week_Day_Night = Week_Day_Night + 1
Case "Weekend Day-Night"
Weekend_Day_Night = Weekend_Day_Night + 1
End Select
End If
data1.Close
data9.MoveNext
Wend
data9.Close
End If
count = count + 1
Wend
repeat_html 1
Print #1, "<TR>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Week Day</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Week Night</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Weekend Day</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Weekend Night</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Week Day-Night</font></strong></TD>"
Print #1, "<TD bgcolor=""#000000""><strong><font face=""times new roman"" color=""#ffffff"" size=""4"">Weekend Day-Night</font></strong></TD>"
Print #1, "</TR>"

repeat_html 1
Print #1, "<TR>"
Print #1, "<TD>", Week_Day, "</TD>"
Print #1, "<TD>", Week_Night, "</TD>"
Print #1, "<TD>", Weekend_Day, "</TD>"
Print #1, "<TD>", Weekend_Night, "</TD>"
Print #1, "<TD>", Week_Day_Night, "</TD>"
Print #1, "<TD>", Weekend_Day_Night, "</TD>"
Print #1, "</TR>"
End If
'////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////

Print #1, " </TABLE>"
Print #1, "</center>"
Print #1, "<P>"
Print #1, "<BR>"
Print #1, "</CENTER>"
Print #1, "</CENTER>"
Print #1, "<center>&copy;1989-2003 UHD Multimedia Services Central Event Managment System/David Martin<center>"
Print #1, "<center>All Rights Reserved<center>"
Print #1, "</BODY></HTML>"
Close #1

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub




Private Sub end1_Click()
On Error GoTo has_error

Dim connection_string As String
connection_string = Form1.connection_string
List1.Clear


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
.Source = "Select * From Schedule_Info Order By [Begin Date] ASC"
End With


With data1
.Open
.MoveFirst
.Find "[Begin Date]= '" & begin.Text & "'"
On Error GoTo gocom
While ((.EOF = False) And (end1.Text <= ![End Date]))
List1.AddItem ![Event]
.MoveNext
Wend
End With
gocom:

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
Dim adn1 As ADODB.Connection
Dim data2 As ADODB.Recordset
Dim data1 As ADODB.Recordset


Set adn = New ADODB.Connection
Set adn1 = New ADODB.Connection
Set data2 = New ADODB.Recordset
Set data1 = New ADODB.Recordset

With adn
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With adn1
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With


With data1
.ActiveConnection = adn
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Schedule_Info Order by [Begin Date] ASC"
End With



With data1
.Open

While .EOF = False
begin.AddItem ![Begin Date]
end1.AddItem ![Begin Date]
.MoveNext
Wend
.Close
End With
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub funk1(temp1 As String, temp2 As String, temp3 As Integer)

On Error GoTo has_error
'******************************************************************
Dim connection_string As String
connection_string = Form1.connection_string
Dim array_n(1 To 50) As String
Dim x As Long


x = 0

Dim adn3 As ADODB.Connection
Dim data10 As ADODB.Recordset

Set adn3 = New ADODB.Connection
Set data10 = New ADODB.Recordset


With adn3
.connectionstring = connection_string
.CursorLocation = adUseClient
.Open
End With

With data10
.ActiveConnection = adn3
.CursorType = adOpenKeyset
.LockType = adLockOptimistic
.Source = "Select * From Check_In Where [Event_ID]= '" & temp1 & "' AND [Event_Date]= #" & temp2 & "# Order by [Event_ID] DESC"
End With
'******************************************************************




If temp3 Mod 2 Then
temp3 = temp3 + 1
End If
Dim sText As String
Dim lTextLength As Long
Dim sChar As String
Dim bASCII As Byte
Dim y As Long
Dim count As Integer
Dim count1 As Integer
count1 = 0


data10.Open
If data10.EOF = False Then
sText = data10![activity]
lTextLength = Len(sText) 'Gets # of chars in sText

For y = 1 To lTextLength 'Loop through string one char at a time
array_n(y) = Mid$(sText, y, 1)
Next y

count = 1
While count < lTextLength
If Val(array_n(count)) = temp3 Then
If count1 = 0 Then
Print #1, "<TD>", data10![Employee_Name], "</TD>"
Print #1, "<TD>", data10![total_hours], "</TD><BR>"
Else
Print #1, "<TD>", data10![Employee_Name], "</TD>"
Print #1, "<TD>", data10![total_hours], "</TD><BR>"
End If
count = count + 2
Else
count = count + 2
End If
Wend
End If
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub




Public Function greater_than(temp1 As String, temp2 As String)

On Error GoTo has_error
Dim first_date As String
Dim second_date As String
Dim month1 As String
Dim day1 As String
Dim year1 As String
Dim month2 As String
Dim day2 As String
Dim year2 As String
Dim found As Boolean
Dim equal As Boolean

found = False
equal = False



first_date = temp1
second_date = temp2

month1 = Month(first_date)
day1 = Month(first_date)
year1 = Month(first_date)

month2 = Month(first_date)
day2 = Month(first_date)
year2 = Month(first_date)

If year1 > year2 Then
    found = True
    ElseIf year1 = year2 Then
        If month1 > month2 Then
            found = True
    ElseIf month1 = month2 Then
        If day1 > day2 Then
            found = True
    ElseIf day1 = day2 Then
        equal = True
        End If
        End If
End If
Exit Function
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Function

Sub repeat_html(count As Integer)
On Error GoTo has_error
Dim x As Integer

For x = 1 To count
Print #1, "<TR>"
Print #1, "<TD></TD>"
Print #1, "<TD></TD>"
Print #1, "</TR>"
Print #1, "<TR>"
Next x

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub


