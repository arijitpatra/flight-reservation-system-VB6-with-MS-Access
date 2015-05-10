VERSION 5.00
Begin VB.Form contacts 
   Caption         =   "Contact"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   -1080
      Picture         =   "contacts.frx":0000
      Top             =   -1080
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
Attribute VB_Name = "contacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub contact_Click()
contacts.Hide
Load contacts
contacts.Show
End Sub

Private Sub exit_Click()
contacts.Hide
End Sub

Private Sub faq_Click()
contacts.Hide
Load faskq
faskq.Show
End Sub

Private Sub home_Click()
contacts.Hide
Load welcome
welcome.Show
End Sub

Private Sub signup_Click()
contacts.Hide
Load sign
sign.Show
End Sub

Private Sub tandc_Click()
contacts.Hide
Load tac
tac.Show
End Sub
