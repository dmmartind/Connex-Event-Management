VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Employee"
   ClientHeight    =   570
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc data2 
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\My Documents\temp.mde;Jet OLEDB:Database Password=51289;"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\My Documents\temp.mde;Jet OLEDB:Database Password=51289;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Employee"
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
   Begin VB.ComboBox Combo1 
      DataField       =   "Employee_Name"
      DataSource      =   "data2"
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Combo1_Change()
Dim strSearch As String
Dim book As Variant
Dim temp1 As Integer
Dim temp2 As Recordset

Set temp2 = Form2.Adodc1.Recordset
Combo1.SetFocus
If Combo1.Text = "" Then
temp1 = MsgBox("Please select an event", vbExclamation, "Event Search")
Else
strSearch = "[event]= '" & Combo1 & "'"
temp2.Find strSearch
Form2.name = temp2![name]
Form2.room = temp2![room]
Form2.time = temp2![time]
Form2.room_admin = temp2![admin]
Form2.group = temp2![group]
Form2.time = temp2![time]
End Sub

Private Sub Form_Activate()
Dim temp1 As Recordset
Dim temp2 As Integer


Set temp1 = Adodc3.Recordset

On Error GoTo handle1

temp1.MoveFirst
While (temp1.EOF <> True)

    Combo1.AddItem temp1![event]
    
Wend

handle1:
temp2 = MsgBox("There are no records in the database to report. Please add a record.", vbInformation)
Unload Me

End Sub

