VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   720
      Top             =   2640
   End
   Begin VB.Frame Frame1 
      Height          =   4155
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   7425
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "�1989-2002 David Martin. All Rights Reserved"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   3480
         Width           =   3615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Software Systems Inc."
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   3240
         Width           =   3735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Mega_Tron"
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   3000
         Width           =   3975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Version 1.0 Beta"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   2280
         Width           =   4095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Central Event Managment System"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   1920
         Width           =   5415
      End
      Begin VB.Image imgLogo 
         Height          =   1545
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   6495
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Timer1_Timer()
Unload Me
Form1.Show vbModal
End Sub
