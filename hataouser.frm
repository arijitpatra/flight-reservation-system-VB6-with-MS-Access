VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form hataouser 
   Caption         =   "Remove User"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   11430
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   3600
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   2400
      Width           =   3975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1095
      Left            =   5880
      Top             =   8400
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
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
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=FRS DB"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=FRS DB"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "signup"
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
   Begin VB.ComboBox removallist 
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "hataouser.frx":0000
      Left            =   6360
      List            =   "hataouser.frx":0002
      TabIndex        =   0
      Text            =   "Select User to Remove"
      Top             =   1680
      Width           =   3975
   End
End
Attribute VB_Name = "hataouser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Adodc1.Recordset.MoveFirst
    While Adodc1.Recordset.EOF = False
        removallist.AddItem (Adodc1.Recordset.Fields(0))
        Adodc1.Recordset.MoveNext
    Wend
End Sub


Private Sub removallist_Change()
If removallist.Text = removallist.List() Then
     Text1.Text = Adodc1.Recordset.Fields(0)
    Text2.Text = Adodc1.Recordset.Fields(1)
    Text3.Text = Adodc1.Recordset.Fields(2)
End Sub







