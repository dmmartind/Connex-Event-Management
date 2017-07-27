VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form12.frx":0000
   ScaleHeight     =   5175
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2520
      OleObjectBlob   =   "Form12.frx":3106
      Top             =   2400
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.Text = Date

End Sub
