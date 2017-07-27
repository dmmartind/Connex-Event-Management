VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form search_e 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Employee"
   ClientHeight    =   585
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin MSAdodcLib.Adodc data30 
      Height          =   330
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   $"search_e.frx":0000
      OLEDBString     =   $"search_e.frx":0087
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Event_Room"
      Caption         =   "data30"
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
   Begin VB.ComboBox Combo1 
      DataField       =   "Employee_Name"
      DataSource      =   "data2"
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "search_e"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit





Private Sub Form_Activate()
data30.Refresh
With data30.Recordset


Do Until data30.Recordset.EOF
    If ![Room_#] <> "" Then
        Combo1.AddItem ![Room_#]
    End If
        .MoveNext
Loop
.Close
End With
End Sub

Private Sub OK_Click()

data30.Refresh
With data30.Recordset
.Find "[Room_#]= '" & Combo1.Text & "'"
If .EOF = False Then
Form6.room.Text = ![Room_#]
Form6.admin_name.Text = ![admin_name]
Form6.phone.Text = ![Admin_Phone]
Form6.email.Text = ![Admin_Email]
Form6.room.Enabled = False
Form6.admin_name.Enabled = False
Form6.phone.Enabled = False
Form6.email.Enabled = False
Form6.edit_record.Enabled = True
Form6.new_record.Enabled = False
Form6.search_record.Enabled = False
Form6.edit_record.Enabled = True
Form6.new_record.Enabled = False
Form6.Add.Enabled = False
Form6.Delete.Enabled = True
Form6.Reset.Enabled = True
Else
MsgBox "No record found!!!", vbExclamation, "Search"
End If
Unload Me
End With
End Sub
