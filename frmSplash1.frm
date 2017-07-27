VERSION 5.00
Begin VB.Form splash1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3225
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   2520
      Top             =   1560
   End
   Begin VB.Image Image1 
      Height          =   3630
      Left            =   -120
      Picture         =   "frmSplash1.frx":000C
      Top             =   -120
      Width           =   5520
   End
End
Attribute VB_Name = "splash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Option Explicit

Private Sub Timer1_Timer()
On Error GoTo has_error
Unload Me
Form1.Show vbModal
Exit Sub
has_error:
MsgBox "Error: Wait a minute!!! You did something we didn't expect you to do. Sorry for the inconvenience this has caused you.", vbExclamation, "Connex Event Management"
Unload Me
End Sub
