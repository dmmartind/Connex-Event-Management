VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Schedule"
   ClientHeight    =   5145
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8040
   LinkTopic       =   "Form2"
   ScaleHeight     =   5145
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox group 
      Height          =   285
      Left            =   5760
      TabIndex        =   13
      Top             =   1920
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3240
      Top             =   2280
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\My Documents\temp.mde;Jet OLEDB:Database Password=51289;"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\My Documents\temp.mde;Jet OLEDB:Database Password=51289;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Schedule_Info"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox name 
      DataField       =   "Event"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox room 
      DataField       =   "Room"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox date 
      DataField       =   "Date"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox room_admin 
      DataField       =   "Administrator"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   5760
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox hour_type 
      DataField       =   "Hour_Type"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   5760
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox time 
      DataField       =   "Time"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   5760
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "view_schedule.frx":0000
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3836
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\My Documents\temp.mde;Jet OLEDB:Database Password=51289;"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\My Documents\temp.mde;Jet OLEDB:Database Password=51289;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Activity"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Schedule"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   2640
      TabIndex        =   16
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label9 
      Caption         =   "Group"
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   3360
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Time"
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Room #"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Name of Event"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Room Admin"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hour Type"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   2520
      Top             =   0
      Width           =   4455
   End
   Begin VB.Menu menu1 
      Caption         =   "Main"
      Begin VB.Menu search_event 
         Caption         =   "Search Event"
         Index           =   0
      End
      Begin VB.Menu E_xit 
         Caption         =   "Exit"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub E_xit_Click(Index As Integer)
Unload Me

End Sub

Private Sub Label8_Click()

End Sub

Private Sub search_event_Click(Index As Integer)
Dialog.Show vbModal
End Sub
