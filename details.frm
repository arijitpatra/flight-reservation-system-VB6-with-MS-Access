VERSION 5.00
Begin VB.Form details 
   Caption         =   "Show"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "home"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "PASSWORD"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "USER ID"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
details.Hide
Load welcome
welcome.Show
End Sub
