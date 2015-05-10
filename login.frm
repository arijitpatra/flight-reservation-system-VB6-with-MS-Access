VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   Caption         =   "Login"
   ClientHeight    =   4965
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1335
      Left            =   1200
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2355
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
   Begin VB.CommandButton forgotpasswordsign 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Forgot Password?"
      Height          =   255
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton loginsign 
      BackColor       =   &H00C0C000&
      Caption         =   "Sign Up"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton loginok 
      BackColor       =   &H0080FFFF&
      Caption         =   "OK"
      Height          =   495
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox userpassword 
      DataSource      =   "Adodc1"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   14760
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox username 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   14760
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Are you a new user?                                                 now!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   3000
      TabIndex        =   7
      Top             =   720
      Width           =   7815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   13680
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   13560
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   0
      Picture         =   "login.frx":0000
      Top             =   0
      Width           =   21000
   End
   Begin VB.Menu home 
      Caption         =   "Home"
   End
   Begin VB.Menu signup 
      Caption         =   "Sign Up"
   End
   Begin VB.Menu faq 
      Caption         =   "FAQ"
   End
   Begin VB.Menu tandc 
      Caption         =   "Terms and Conditions"
   End
   Begin VB.Menu contact 
      Caption         =   "Contact"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub contact_Click()
login.Hide
Load contacts
contacts.Show
End Sub

Private Sub exit_Click()
login.Hide
End Sub

Private Sub faq_Click()
login.Hide
Load faskq
faskq.Show
End Sub


Private Sub forgotpasswordsign_Click()
MsgBox ("Sorry,we don't remember your passwords! Create a new account by clicking Sign Up or login as Guest.")
login.Hide
Load welcome
welcome.Show
End Sub

Private Sub Form_Load()
username.Text = ""
userpassword.Text = ""
End Sub

Private Sub home_Click()
login.Hide
Load welcome
welcome.Show

End Sub

Private Sub loginok_Click()
If username.Text = "" Then
MsgBox ("Enter Username")
ElseIf userpassword.Text = "" Then
MsgBox ("Enter password")
End If
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find "username='" & username.Text & "'"

If Adodc1.Recordset.EOF Then
     username.Text = ""
     userpassword.Text = ""
     MsgBox "Username Not Found"
     
 Else
     If Adodc1.Recordset.Fields(4) = userpassword.Text Then
     MsgBox "You Are Now Logged In. Happy Flying!"
     login.Hide
     Load journeydetails
     journeydetails.Show
     Eticket.bookingagent.Text = username.Text
     Else
     MsgBox "Password Don't Match"
     userpassword.Text = ""
     End If
End If

End Sub


Private Sub loginsign_Click()
login.Hide
Load sign
sign.Show
End Sub

Private Sub signup_Click()
login.Hide
Load sign
sign.Show
End Sub

Private Sub tandc_Click()
login.Hide
Load tac
tac.Show
End Sub
