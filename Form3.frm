VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form check_in 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Clock-In"
   ClientHeight    =   3945
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6915
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form3.frx":0000
      TabIndex        =   12
      Top             =   2760
      Width           =   855
   End
   Begin MSMask.MaskEdBox event_date 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   2760
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   0
      Picture         =   "Form3.frx":0072
      ScaleHeight     =   435
      ScaleWidth      =   6795
      TabIndex        =   10
      Top             =   0
      Width           =   6855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   4440
      OleObjectBlob   =   "Form3.frx":301C
      TabIndex        =   9
      Top             =   2280
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   4440
      OleObjectBlob   =   "Form3.frx":3082
      TabIndex        =   8
      Top             =   1800
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form3.frx":30E8
      TabIndex        =   7
      Top             =   2280
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form3.frx":3150
      TabIndex        =   6
      Top             =   1800
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "Form3.frx":31B6
      Top             =   840
   End
   Begin VB.TextBox name1 
      DataField       =   "Employee_Name"
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1800
      Width           =   2775
   End
   Begin VB.ComboBox event1 
      DataField       =   "Event_ID"
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton Reset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Clock_in 
      Caption         =   "Clock-In"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox time1 
      DataField       =   "In_Time"
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox date1 
      DataField       =   "Date"
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Menu file_1 
      Caption         =   "File"
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "check_in"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Dim record_control2 As Integer
Dim opps As Integer


Private Sub Clock_in_Click()
'************************************************************************************8
'***********************************************************************************
On Error GoTo has_error

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
.Source = "Select * From Check_in Order by [Event_ID]"
End With



'***********************************************************************
'************************************************************************
If (record_control2 = 0) Then
With data1
.Open
.AddNew
.Fields(1) = event1.Text
.Fields(2) = name1.Text
.Fields(3) = date1.Text
.Fields(4) = time1.Text
.Fields(7) = event_date.Text
.Update
.Close
End With
End If
adn.Close

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me


End Sub

Private Sub event1_Click()
On Error GoTo has_error
Dim adn As ADODB.Connection
Dim data1 As ADODB.Recordset

Dim connection_string As String
connection_string = Form1.connection_string

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
.Source = "Select * From Schedule_Info where [Event]= '" & event1.Text & "'"
End With
'/////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
With data1
.Open
If .EOF = False Then
event_date.Text = ![Begin Date]
End If
.Close
End With
adn.Close

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me


End Sub

Private Sub Reset_Click()
On Error GoTo has_error
name1.Locked = True
name1.Text = Form1.user_name1
date1.Text = Date
time1.Text = Time
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me


End Sub
Private Sub exit_Click()
On Error GoTo has_error
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
'*************************************************************************
connection_string = Form1.connection_string



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
.Source = "Select * From Schedule_Info Order by [Event]"
End With


'*************************************************************************
name1.Locked = True
name1.Text = Form1.user_name1
date1.Text = Date
time1.Text = Time


With data2
.Open

Do Until data2.EOF
    If ![Event] <> "" Then
        event1.AddItem ![Event]
    End If
        .MoveNext
Loop
.Close
End With
name1.Text = Form1.user_name1
adn1.Close
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub form_queryUnload(cancel As Integer, unloadmode As Integer)
'*************************************************************************
On Error GoTo has_error
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
.Source = "Select * From Check_In Order by [Event_ID]"
End With


'*************************************************************************






With data1
.Open
Dim temp As Integer
If opps = 1 Then
temp = MsgBox("Do you really want to quit?", vbYesNo, "Connex Event Management")
Select Case temp
Case 6
    .cancel
    Unload Me
Case 7
    cancel = 1
End Select
Else
Unload Me
End If
.Close
End With
adn.Close
Exit Sub

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me






End Sub


