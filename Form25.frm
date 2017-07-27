VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   LinkTopic       =   "Form10"
   ScaleHeight     =   5505
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6240
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   2160
      OleObjectBlob   =   "Form25.frx":0000
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   5400
      OleObjectBlob   =   "Form25.frx":0074
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Form25.frx":00DC
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   1920
      Width           =   2295
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2003 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Private Sub Form_Load()
Dim connection_string As Integer

connection_string = Form1.connection_string

Dim adn1 As ADODB.Connection
Dim data1 As ADODB.Recordset

End Sub
