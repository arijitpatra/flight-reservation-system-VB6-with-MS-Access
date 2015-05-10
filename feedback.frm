VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form feedback 
   BackColor       =   &H00004000&
   Caption         =   "Feedback"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11130
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton showform 
      BackColor       =   &H0000FFFF&
      Caption         =   "Show me the feedback form"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   3495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   7320
      TabIndex        =   8
      Top             =   3120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6800
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1455
      Left            =   4920
      Top             =   9600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2566
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
      RecordSource    =   "feedback"
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
   Begin VB.CommandButton submitfeedback 
      BackColor       =   &H0000FFFF&
      Caption         =   "SUBMIT FEEDBACK :)"
      Enabled         =   0   'False
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   2055
   End
   Begin VB.ComboBox rating 
      DataField       =   "Rating"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "feedback.frx":0000
      Left            =   4920
      List            =   "feedback.frx":0016
      TabIndex        =   2
      Text            =   "---Rate Our Service---"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox feedname 
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox feed 
      DataField       =   "Feedback"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FEEDBACK"
      BeginProperty Font 
         Name            =   "Aller"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   8760
      TabIndex        =   10
      Top             =   840
      Width           =   7935
   End
   Begin VB.Label rate 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Our Service"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label feedlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Thank You for using our service. Please help us making your experience better by giving your feedback. It takes just a minute... "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   5760
      TabIndex        =   6
      Top             =   1560
      Width           =   8175
   End
   Begin VB.Label yourname 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name : "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label yourfeedback 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Feedback : "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Menu hme 
      Caption         =   "Home"
   End
   Begin VB.Menu sgn 
      Caption         =   "Sign Up"
   End
   Begin VB.Menu fq 
      Caption         =   "FAQ"
   End
   Begin VB.Menu tc 
      Caption         =   "Terms and Conditions "
   End
   Begin VB.Menu cntct 
      Caption         =   "Contact"
   End
   Begin VB.Menu ext 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "feedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPatient As New ADODB.Recordset

Private Sub cntct_Click()
feedback.Hide
Load contacts
contacts.Show
End Sub

Private Sub ext_Click()
feedback.Hide
End Sub

Private Sub feedbackhome_Click()
feedback.Hide
Load welcome
welcome.Show
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
 Set res = New ADODB.Recordset
 con.CursorLocation = adUseClient
 con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\arijit\FRS1\database\FRS DB.mdb;Persist Security Info=False"
 con.Open
 
 res.Open "select * from feedback", con, adOpenStatic, adLockOptimistic

Set DataGrid1.DataSource = res

 DataGrid1.Refresh
End Sub

Private Sub fq_Click()
feedback.Hide
Load faskq
faskq.Show
End Sub

Private Sub hme_Click()
feedback.Hide
Load welcome
welcome.Show
End Sub

Private Sub sgn_Click()
feedback.Hide
Load sign
sign.Show
End Sub

Private Sub showform_Click()
Adodc1.Recordset.AddNew
feedlabel.Visible = False
showform.Visible = False
feed.Visible = True
rating.Visible = True
rate.Visible = True
feedname.Visible = True
yourfeedback.Visible = True
yourname.Visible = True
submitfeedback.Enabled = True
End Sub

Private Sub submitfeedback_Click()
    Adodc1.Recordset.Fields(0) = feed.Text
    Adodc1.Recordset.Fields(1) = rating.Text
    Adodc1.Recordset.Fields(2) = feedname.Text
    Adodc1.Recordset.Update
    feed.Visible = False
rating.Visible = False
feedname.Visible = False
yourfeedback.Visible = False
yourname.Visible = False
rate.Visible = False
submitfeedback.Visible = False
    MsgBox "Thank You"
    feedback.Hide
Load welcome
welcome.Show
    
End Sub

Private Sub tc_Click()
feedback.Hide
Load tac
tac.Show
End Sub
