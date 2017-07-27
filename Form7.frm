VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form7 
   Caption         =   "NeMos Backup Setup"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel label2 
      Height          =   255
      Left            =   2400
      OleObjectBlob   =   "Form7.frx":0000
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form7.frx":005E
      Left            =   1080
      List            =   "Form7.frx":00A1
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   2760
      Picture         =   "Form7.frx":00F3
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "Form7.frx":0265
      Top             =   2880
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form7.frx":4B064
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   480
      Picture         =   "Form7.frx":4B0E4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   720
      OleObjectBlob   =   "Form7.frx":4B616
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   1440
      TabIndex        =   0
      Top             =   2880
      Width           =   1695
      Begin MSComctlLib.Slider Slider1 
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         _Version        =   393216
         Max             =   1
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "Form7.frx":4B6B6
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Form7.frx":4B718
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Label Label1 
      Caption         =   "©2003 David Martin. All Rights Reserved"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   3840
      Width           =   3015
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Public directory1 As String
Public voice As Integer
Public index As Integer



Private Sub Combo1_Click()

On Error GoTo has_error
Select Case Combo1.ListIndex
Case 0
    index = 1
    label2.Caption = "Minute"
Case 1
    index = 5
    label2.Caption = "Minutes"
Case 2
    index = 10
    label2.Caption = "Minutes"
Case 3
    index = 20
    label2.Caption = "Minutes"
Case 4
    index = 30
    label2.Caption = "Minutes"
Case 5
    index = 40
    label2.Caption = "Minutes"
Case 6
    index = 50
    label2.Caption = "Minutes"
Case 7
    index = 60
    label2.Caption = "Minutes"
Case 8
    index = 70
    label2.Caption = "Minutes"
Case 9
    index = 80
    label2.Caption = "Minutes"
Case 10
    index = 90
    label2.Caption = "Minutes"
Case 11
    index = 100
    label2.Caption = "Minutes"
Case 12
    index = 110
    label2.Caption = "Minutes"
Case 13
    index = 120
    label2.Caption = "Minutes"
Case 14
    index = 1440
    label2.Caption = "Day"
Case 15
    index = 2880
    label2.Caption = "Days"
Case 16
    index = 4320
    label2.Caption = "Days"
Case 17
    index = 5760
    label2.Caption = "Days"
Case 18
    index = 7200
    label2.Caption = "Days"
Case 19
    index = 8640
    label2.Caption = "Days"
Case 20
    index = 10080
    label2.Caption = "Week"
Case Else
    MsgBox "How much time between backups???", vbQuestion, "Connex Event Management"
End Select
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub Command1_Click()
On Error GoTo has_error
Form8.Show vbModal
Text1.Text = directory1
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo has_error
Dim adn1 As ADODB.Connection
Dim data1 As ADODB.Recordset
Dim connectionstring As String

Set adn1 = New ADODB.Connection
Set data1 = New ADODB.Recordset

connectionstring = Form1.connection_string

With adn1
.connectionstring = connectionstring
.CursorLocation = adUseClient
.Open
End With

With data1
.ActiveConnection = adn1
.CursorType = adOpenDynamic
.LockType = adLockOptimistic
.Source = "SELECT * FROM backup"
End With

On Error GoTo heck

With data1
.Open
While .EOF = False
.Delete
.MoveNext
Wend



.AddNew
.Fields(1) = index
.Fields(2) = Text1.Text
.Fields(3) = Slider1.Value
.Update
.Close
End With
adn1.Close
If Slider1.Value = 0 And voice = 1 Then
Form1.NeMos_deac1
End If
voice = 0
Unload Me
Exit Sub


heck:
If Text1.Text = "" Then
MsgBox "Where do you want me save it?? You need to specify the save directory", vbExclamation, "Connex Event Management"
End If
If Text1.Text <> "" Then
MsgBox "The directory is TOO LONG!!!!! Please pick a shorter one. Sorry, :-( ", vbExclamation, "Connex Event Management"
End If
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo has_error
Skin1.ApplySkin Me.hwnd


Dim adn1 As ADODB.Connection
Dim data1 As ADODB.Recordset

Set adn1 = New ADODB.Connection
Set data1 = New ADODB.Recordset

With adn1
.connectionstring = Form1.connection_string
.CursorLocation = adUseClient
.Open
End With

With data1
.ActiveConnection = adn1
.CursorType = adOpenDynamic
.Source = "SELECT * FROM backup"
End With



With data1
.Open
If .EOF = False Then
Combo1.Text = ![Wait Time]
Text1.Text = ![Backup Directory]
Slider1.Value = ![Power Switch]

If Slider1.Value = 1 Then
voice = 1
End If

End If
.Close
End With
adn1.Close

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

