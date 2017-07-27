VERSION 5.00
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "VBSKFREE.OCX"
Begin VB.Form employee 
   Caption         =   "Form5"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2040
   LinkTopic       =   "Form5"
   ScaleHeight     =   1155
   ScaleWidth      =   2040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   2520
      TabIndex        =   0
      Top             =   3000
      Width           =   2055
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   1200
      Top             =   1560
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
End
Attribute VB_Name = "employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
