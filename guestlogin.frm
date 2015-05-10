VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form guestlogin 
   Caption         =   "Guest Login"
   ClientHeight    =   4605
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8760
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H0000FFFF&
      Caption         =   "HERE"
      BeginProperty Font 
         Name            =   "Nokia Standard Multiscript"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1335
      Left            =   1320
      Top             =   3120
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "guest"
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
   Begin VB.CommandButton guestok 
      BackColor       =   &H000000FF&
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   495
      Left            =   14640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox phonenum 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   14160
      TabIndex        =   3
      Text            =   "+91"
      Top             =   1560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox emailid 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   14160
      TabIndex        =   2
      Text            =   "@"
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "to unlock the fields"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5880
      TabIndex        =   7
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Click "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   5760
      TabIndex        =   6
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone :"
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
      Height          =   375
      Left            =   13320
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail ID :"
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
      Height          =   375
      Left            =   13080
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   0
      Picture         =   "guestlogin.frx":0000
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
Attribute VB_Name = "guestlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Adodc1.Recordset.AddNew
cmdAdd.Caption = "Great!"
emailid.Visible = True
phonenum.Visible = True
guestok.Enabled = True
guestok.BackColor = &HC000&
Label3.Visible = False
Label4.Visible = False
cmdAdd.Enabled = False
End Sub

Private Sub contact_Click()
guestlogin.Hide
Load contacts
contacts.Show
End Sub

Private Sub exit_Click()
guestlogin.Hide
End Sub

Private Sub faq_Click()
guestlogin.Hide
Load faskq
faskq.Show
End Sub

Private Sub guestok_Click()
If (emailid.Text = "") Then
MsgBox ("User Name please :)")
ElseIf (phonenum.Text = "") Then
MsgBox ("Phone No. please :)")
Else
    Adodc1.Recordset.Fields(0) = emailid.Text
    Adodc1.Recordset.Fields(1) = phonenum.Text
    Adodc1.Recordset.Update
guestlogin.Hide
Load journeydetails
journeydetails.Show
Eticket.bookingagent.Text = guestlogin.emailid.Text
End If
End Sub

Private Sub home_Click()
guestlogin.Hide
Load welcome
welcome.Show
End Sub

Private Sub signup_Click()
guestlogin.Hide
Load sign
sign.Show
End Sub

Private Sub tandc_Click()
guestlogin.Hide
Load tac
tac.Show
End Sub
