VERSION 5.00
Begin VB.Form welcome 
   BackColor       =   &H80000002&
   Caption         =   "Welcome"
   ClientHeight    =   8190
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11670
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   240
   End
   Begin VB.CommandButton guest 
      BackColor       =   &H0080FFFF&
      Caption         =   "Guest"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton user 
      BackColor       =   &H0080FFFF&
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton admin 
      BackColor       =   &H0080FFFF&
      Caption         =   "Admin"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1755
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "© FRS 2014 | Ankita Gupta - Anurag Ghosh - Arijeet Saha -Arijit Patra | CSE | 2nd Yr | NITMAS "
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
      Height          =   855
      Left            =   5160
      TabIndex        =   7
      Top             =   10080
      Width           =   9255
   End
   Begin VB.Image Image2 
      Height          =   8175
      Left            =   0
      Top             =   0
      Width           =   11535
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   15480
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Flight Reservation System"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   45
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "                    It's Fast | It's Cheap | It's Easy"
      BeginProperty Font 
         Name            =   "Nokia Standard Multiscript"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   14040
      TabIndex        =   5
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Fly anywhere in India"
      BeginProperty Font 
         Name            =   "Nokia Standard Multiscript"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   14280
      TabIndex        =   4
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   13320
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400040&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   14400
      Top             =   4920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label preference 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select preference:"
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
      Left            =   10920
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   0
      Top             =   -120
      Width           =   21000
   End
   Begin VB.Image Image3 
      Height          =   12225
      Left            =   -600
      Picture         =   "welcome.frx":0000
      Top             =   -960
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
   Begin VB.Menu fdbck 
      Caption         =   "Feedback"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub admin_Click()
welcome.Hide
Load adminlogin
adminlogin.Show
End Sub


Private Sub contact_Click()
welcome.Hide
Load contacts
contacts.Show
End Sub

Private Sub exit_Click()
welcome.Hide
End Sub

Private Sub faq_Click()
welcome.Hide
Load faskq
faskq.Show
End Sub

Private Sub fdbck_Click()
welcome.Hide
Load feedback
feedback.Show
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Timer1.Interval = 1000
End Sub

Private Sub guest_Click()
welcome.Hide
Load guestlogin
guestlogin.Show
End Sub

Private Sub home_Click()
welcome.Hide
Load welcome
welcome.Show

End Sub

Private Sub signup_Click()
welcome.Hide
Load sign
sign.Show
End Sub


Private Sub tandc_Click()
welcome.Hide
Load tac
tac.Show
End Sub

Private Sub Timer1_Timer()

If (Shape1.FillStyle = 1) Then
Shape1.FillStyle = 0
Shape2.FillStyle = 1
Shape3.FillStyle = 0
Label1.Visible = True
Label2.Visible = True
preference.Visible = True

ElseIf (Shape2.FillStyle = 1) Then
Shape1.FillStyle = 1
Shape2.FillStyle = 0
Shape3.FillStyle = 1
Label1.Visible = True
Label2.Visible = True
preference.Visible = False


Else

Shape1.FillStyle = 1
Shape2.FillStyle = 1
Shape3.FillStyle = 1
Label1.Visible = True
Label2.Visible = True
preference.Visible = True

End If






End Sub

Private Sub user_Click()
welcome.Hide
Load login
login.Show
End Sub
