VERSION 5.00
Begin VB.Form journeydetails 
   BackColor       =   &H8000000B&
   Caption         =   "Journey Details"
   ClientHeight    =   5400
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   11130
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame journeyframe 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Journey Details"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   4095
      Left            =   10200
      TabIndex        =   0
      Top             =   720
      Width           =   7575
      Begin VB.ComboBox specialconcession 
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
         ItemData        =   "journeydetails.frx":0000
         Left            =   4920
         List            =   "journeydetails.frx":0016
         TabIndex        =   15
         Text            =   "Spl. Concess."
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton search 
         BackColor       =   &H0000FFFF&
         Caption         =   "I am done! Search for flights!"
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
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3240
         Width           =   3495
      End
      Begin VB.ComboBox seniorcitizen 
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
         ItemData        =   "journeydetails.frx":002C
         Left            =   3600
         List            =   "journeydetails.frx":0042
         TabIndex        =   13
         Text            =   "Sr. Ctzn."
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox child 
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
         ItemData        =   "journeydetails.frx":0058
         Left            =   2400
         List            =   "journeydetails.frx":006E
         TabIndex        =   12
         Text            =   "Child"
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox adults 
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
         ItemData        =   "journeydetails.frx":0084
         Left            =   1200
         List            =   "journeydetails.frx":009A
         TabIndex        =   11
         Text            =   "Adults"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox dateofreturn 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox dateofjourney 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   1680
         Width           =   1695
      End
      Begin VB.OptionButton oneway 
         BackColor       =   &H00FFFFFF&
         Caption         =   "One Way"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton roundway 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Round Way"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox from 
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
         ItemData        =   "journeydetails.frx":00B0
         Left            =   1920
         List            =   "journeydetails.frx":00C0
         TabIndex        =   2
         Text            =   "---From---"
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox to 
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
         ItemData        =   "journeydetails.frx":00E5
         Left            =   3840
         List            =   "journeydetails.frx":00F5
         TabIndex        =   1
         Text            =   "---To---"
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "      (DD/MM/YYYY)"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   5400
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(DD/MM/YYYY)"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "     Date of Return"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   495
         Left            =   3840
         TabIndex        =   8
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "      Date of Journey"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   0
      Picture         =   "journeydetails.frx":011A
      Top             =   0
      Width           =   21000
   End
   Begin VB.Menu home 
      Caption         =   "Home"
   End
   Begin VB.Menu signup 
      Caption         =   "Sign up"
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
   Begin VB.Menu sgno 
      Caption         =   "Sign Out"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "journeydetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub contact_Click()
journeydetails.Hide
Load contacts
contacts.Show
End Sub


Private Sub exit_Click()
journeydetails.Hide
End Sub

Private Sub faq_Click()
journeydetails.Hide
Load faskq
faskq.Show
End Sub

Private Sub Form_Load()
Label2.Visible = False
Label4.Visible = False
dateofreturn.Visible = False
oneway.Value = True
End Sub

Private Sub home_Click()
journeydetails.Hide
Load welcome
welcome.Show
End Sub

Private Sub oneway_Click()
Label2.Visible = False
Label4.Visible = False
dateofreturn.Visible = False
End Sub

Private Sub roundway_Click()
If roundway.Value = True Then
Label2.Visible = True
Label4.Visible = True
dateofreturn.Visible = True
End If
searchresult.flights.Visible = False
searchresult.return.Visible = True
Eticket.returnflightnum.Visible = True
Eticket.returndeparture.Visible = True
Eticket.returnarrival.Visible = True
Eticket.returnstatus.Visible = True
Eticket.returnflightinfo.Visible = True
End Sub

Private Sub search_Click()
MsgBox "Please wait while we search flights for you!"
'Eticket.departing.Text = journeydetails.dateofjourney.Text
journeydetails.Hide
searchresult.Show
End Sub

Private Sub sgno_Click()
journeydetails.Hide
Load welcome
welcome.Show
End Sub

Private Sub signup_Click()
journeydetails.Hide
Load sign
sign.Show
End Sub

Private Sub tandc_Click()
journeydetails.Hide
Load tac
tac.Show
End Sub
