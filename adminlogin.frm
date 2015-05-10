VERSION 5.00
Begin VB.Form adminlogin 
   Caption         =   "Admin Login"
   ClientHeight    =   4815
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   8025
   WindowState     =   2  'Maximized
   Begin VB.CommandButton adminok 
      BackColor       =   &H0000FFFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Nokia Standard Multiscript"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ADMIN LOGIN"
      BeginProperty Font 
         Name            =   "Aller Light"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11280
      TabIndex        =   6
      Top             =   9240
      Width           =   7095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"adminlogin.frx":0000
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   840
      TabIndex        =   5
      Top             =   4800
      Width           =   8895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
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
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID :"
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
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   -1440
      Picture         =   "adminlogin.frx":00AD
      Top             =   -840
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
Attribute VB_Name = "adminlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adminok_Click()
If Text1.Text = "admin" Then
If Text2.Text = "password" Then
MsgBox ("Authenticated!")
adminlogin.Hide
Load admincontrol
admincontrol.Show
End If
Else
MsgBox ("Error!")
End If
Text1.Text = " "
Text2.Text = " "
End Sub

Private Sub contact_Click()
adminlogin.Hide
Load contacts
contacts.Show

End Sub

Private Sub exit_Click()
adminlogin.Hide
End Sub

Private Sub faq_Click()
adminlogin.Hide
Load faskq
faskq.Show
End Sub

Private Sub home_Click()
adminlogin.Hide
Load welcome
welcome.Show
End Sub

Private Sub signup_Click()
adminlogin.Hide
Load sign
sign.Show

End Sub

Private Sub tandc_Click()
adminlogin.Hide
Load tac
tac.Show
End Sub
