VERSION 5.00
Begin VB.Form faskq 
   Caption         =   "FAQ"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   -1800
      Picture         =   "faskq.frx":0000
      Top             =   -600
      Width           =   21000
   End
   Begin VB.Menu hme 
      Caption         =   "Home"
   End
   Begin VB.Menu sgup 
      Caption         =   "Sign Up"
   End
   Begin VB.Menu fq 
      Caption         =   "FAQ"
   End
   Begin VB.Menu tnc 
      Caption         =   "Terms and Conditions"
   End
   Begin VB.Menu cntct 
      Caption         =   "Contact"
   End
   Begin VB.Menu ext 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "faskq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cntct_Click()
faskq.Hide
Load contacts
contacts.Show
End Sub

Private Sub ext_Click()
faskq.Hide
End Sub

Private Sub fq_Click()
faskq.Hide
Load faskq
faskq.Show
End Sub

Private Sub hme_Click()
faskq.Hide
Load welcome
welcome.Show
End Sub

Private Sub sgup_Click()
faskq.Hide
Load sign
sign.Show
End Sub

Private Sub tnc_Click()
faskq.Hide
Load tac
tac.Show
End Sub
