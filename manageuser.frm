VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form manageuser 
   Caption         =   "See Users"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "DELETE THIS USER"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   5415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1215
      Left            =   4680
      Top             =   9960
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2143
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\arijit\FRS1\database\FRS DB.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\arijit\FRS1\database\FRS DB.mdb;Persist Security Info=False"
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
   Begin VB.TextBox namesearch 
      Height          =   525
      Left            =   3120
      TabIndex        =   1
      Top             =   1800
      Width           =   7215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5895
      Left            =   3120
      TabIndex        =   0
      Top             =   2640
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   10398
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Username You Want To Remove : "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   0
      Picture         =   "manageuser.frx":0000
      Top             =   0
      Width           =   21000
   End
   Begin VB.Menu adhome 
      Caption         =   "Admin Home"
   End
   Begin VB.Menu mngusr 
      Caption         =   "Manage Users"
   End
   Begin VB.Menu seeusr 
      Caption         =   "See Users"
   End
   Begin VB.Menu logo 
      Caption         =   "Log Out"
   End
End
Attribute VB_Name = "manageuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rsPatient As New ADODB.Recordset
Dim str1, str2 As String
Public str3 As String


Private Sub adhome_Click()
manageuser.Hide
Load admincontrol
admincontrol.Show
End Sub

Private Sub Command1_Click()
manageuser.Hide
Load welcome
welcome.Show

End Sub

Private Sub Command2_Click()
Set rsPatient = New ADODB.Recordset
    If rsPatient.State = adStateOpen Then rsPatient.Close
    rsPatient.Open "Select * from signup", con, adOpenDynamic, adLockOptimistic
    If rsPatient.BOF Then
    MsgBox "Empty Users!"
    Else
       rsPatient.Delete
End If
MsgBox "Record Deleted"
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
 Set res = New ADODB.Recordset
 con.CursorLocation = adUseClient
 con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\arijit\FRS1\database\FRS DB.mdb;Persist Security Info=False"
 con.Open
 
 res.Open "select * from signup", con, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = res
 DataGrid1.Refresh
End Sub

Private Sub logo_Click()
manageuser.Hide
Load welcome
welcome.Show
End Sub

Private Sub mngusr_Click()
manageuser.Hide
Load manageuser
manageuser.Show
End Sub

Private Sub namesearch_Change()

Set rsPatient = New ADODB.Recordset
    If rsPatient.State = adStateOpen Then rsPatient.Close
    rsPatient.Open "Select * from signup", con, adOpenDynamic, adLockOptimistic
' txtsearch = " " Then
'Call Form_Load
'Else

If namesearch.Text = "" Then
Call Form_Load
Me.Show
Else
rsPatient.Filter = "username LIKE '" & Me.namesearch.Text & "*'"
Set DataGrid1.DataSource = rsPatient
End If
End Sub




Private Sub seeusr_Click()
manageuser.Hide
Load seeuser
seeuser.Show
End Sub
