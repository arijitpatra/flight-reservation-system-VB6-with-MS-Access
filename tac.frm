VERSION 5.00
Begin VB.Form tac 
   Caption         =   "Terms and Conditions"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   -840
      Picture         =   "tac.frx":0000
      Top             =   -480
      Width           =   21000
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
      Caption         =   "Terms and Conditions"
   End
   Begin VB.Menu cntct 
      Caption         =   "Contact"
   End
   Begin VB.Menu ext 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "tac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cntct_Click()
tac.Hide
Load contacts
contacts.Show
End Sub

Private Sub ext_Click()
tac.Hide
End Sub

Private Sub fq_Click()
tac.Hide
Load faskq
faskq.Show
End Sub

Private Sub hme_Click()
tac.Hide
Load welcome
welcome.Show
End Sub

Private Sub sgn_Click()
tac.Hide
Load sign
sign.Show
End Sub

Private Sub tc_Click()
tac.Hide
Load tac
tac.Show
End Sub
