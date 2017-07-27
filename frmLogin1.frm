VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmLogin1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unlock"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   495
      Left            =   0
      OleObjectBlob   =   "frmLogin1.frx":0000
      TabIndex        =   3
      Top             =   0
      Width           =   3735
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3240
      OleObjectBlob   =   "frmLogin1.frx":0100
      Top             =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   1
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   2
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   2325
   End
End
Attribute VB_Name = "frmLogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()

On Error GoTo has_error
    LoginSucceeded = False
    Me.Hide
    
    Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me

End Sub

Private Sub cmdOK_Click()

On Error GoTo has_error
    
    If txtPassword = "889D-CDA2-8336-F41" Then
        LoginSucceeded = True
        Me.Hide
       Form1.LoginSucceeded = True
       Unload Me
    Else
        MsgBox "Invalid License Number, try again!", , "Connex Event Management"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
    
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
