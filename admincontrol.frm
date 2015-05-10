VERSION 5.00
Begin VB.Form admincontrol 
   Caption         =   "Admin Control"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   2160
      TabIndex        =   0
      Top             =   1200
      Width           =   7215
      Begin VB.ComboBox admintodo 
         Height          =   315
         ItemData        =   "admincontrol.frx":0000
         Left            =   4080
         List            =   "admincontrol.frx":0002
         TabIndex        =   2
         Text            =   "Select"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "What Do You Want To Do?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3615
      End
   End
   Begin VB.Menu adminhome 
      Caption         =   "Admin &Home"
   End
   Begin VB.Menu signout 
      Caption         =   "&Sign Out"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "admincontrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub adminhome_Click()
admincontrol.Hide
Load admincontrol
admincontrol.Show
End Sub

Private Sub contact_Click()
admincontrol.Hide
Load contacts
contacts.Show
End Sub
Private Sub admintodo_Click()
If admintodo.Text = "View Reservation Details" Then
admincontrol.Hide
viewreservation.Show
ElseIf admintodo.Text = "See Users" Then
admincontrol.Hide
seeusers.Show
End If
End Sub

Private Sub exit_Click()
admincontrol.Hide
End Sub

Private Sub faq_Click()
admincontrol.Hide
Load faqs
faqs.Show
End Sub


Private Sub signup_Click()
admincontrol.Hide
Load sign
sign.Show
End Sub

Private Sub tandc_Click()
admincontrol.Hide
Load tac
tac.Show
End Sub

Private Sub Form_Load()
With admintodo
.AddItem ("View Reservation Details")
.AddItem ("See Users")
End With

End Sub

Private Sub signout_Click()
admincontrol.Hide
Load welcome
welcome.Show
End Sub
