VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form sign 
   BackColor       =   &H80000011&
   Caption         =   "Sign Up"
   ClientHeight    =   5310
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   1200
      Top             =   8760
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Please Fill In The Details"
      ForeColor       =   &H8000000E&
      Height          =   4575
      Left            =   7320
      TabIndex        =   0
      Top             =   3480
      Width           =   5175
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0000C000&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3960
         Width           =   5175
      End
      Begin VB.TextBox confirm 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   2760
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox password 
         DataField       =   "password"
         DataSource      =   "Adodc1"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   2400
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox username 
         DataField       =   "username"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox numphone 
         DataField       =   "phone"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox email 
         DataField       =   "email id"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox txtname 
         DataField       =   "name"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CommandButton done 
         BackColor       =   &H000000FF&
         Caption         =   "DONE"
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password :"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Username :"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone :"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail ID :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "FRS"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2415
      Left            =   15360
      TabIndex        =   16
      Top             =   8880
      Width           =   4815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome! Happy Signing In...   :)"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1455
      Left            =   360
      TabIndex        =   15
      Top             =   1680
      Width           =   13455
   End
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   0
      Picture         =   "sign.frx":0000
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
Attribute VB_Name = "sign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
Adodc1.Recordset.AddNew
cmdAdd.Visible = False
txtname.Visible = True
email.Visible = True
numphone.Visible = True
username.Visible = True
password.Visible = True
confirm.Visible = True
txtname.SetFocus
done.Visible = True
done.Enabled = True
End Sub

Private Sub contact_Click()
sign.Hide
Load contacts
contacts.Show

End Sub

Private Sub done_Click()
If (txtname.Text = "") Then
MsgBox ("Name missing :)")
ElseIf (email.Text = "") Then
MsgBox ("Email missing :)")
ElseIf (numphone.Text = "") Then
MsgBox ("Phone no. missing :)")
ElseIf (username.Text = "") Then
MsgBox ("No Username :)")
ElseIf (password.Text = "") Then
MsgBox ("Password missing :)")
ElseIf (confirm.Text = "") Then
MsgBox ("Renter password :)")
Else
    Adodc1.Recordset.Fields(0) = txtname.Text
    Adodc1.Recordset.Fields(1) = email.Text
    Adodc1.Recordset.Fields(2) = numphone.Text
    Adodc1.Recordset.Fields(3) = username.Text
    Adodc1.Recordset.Fields(4) = password.Text
End If

If password.Text = confirm.Text Then
done.BackColor = &HC000&
MsgBox ("Passwords Match!")
Adodc1.Recordset.Update

MsgBox ("Congrats! You are now successfully signed up")
sign.Hide
Load login
login.Show
Else
MsgBox ("Sorry,passwords don't match")
password.Text = ""
confirm.Text = ""
End If
End Sub

Private Sub exit_Click()
sign.Hide
End Sub

Private Sub faq_Click()
sign.Hide
Load faskq
faskq.Show
End Sub

Private Sub Form_Load()
done.Enabled = False
End Sub

Private Sub home_Click()
sign.Hide
Load welcome
welcome.Show

End Sub

Private Sub signup_Click()
sign.Hide
Load sign
sign.Show
End Sub

Private Sub tandc_Click()
sign.Hide
Load tac
tac.Show
End Sub
