VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Settings1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Directory"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2880
      OleObjectBlob   =   "Form20.frx":0000
      Top             =   1320
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   4680
      OleObjectBlob   =   "Form20.frx":2A5E9
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "Form20.frx":2A653
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   4080
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   3735
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox add_text 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   4215
   End
End
Attribute VB_Name = "Settings1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Private Sub cmdCancel_Click()
On Error GoTo has_error
Form1.setting_temp = 0
Unload Me

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub Dir1_Change()
On Error GoTo has_error
add_text.Text = Dir1.Path

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

Private Sub Drive1_Change()
On Error GoTo has_error
Dim temp As String
temp = Drive1.drive
Dir1.Path = temp

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub




Private Sub Form_Load()
On Error GoTo has_error
Skin1.ApplySkin Me.hwnd

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub OK_Click()
On Error GoTo has_error
Form1.setting_temp = 1
Form1.directory1 = add_text.Text
Unload Me

Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub

