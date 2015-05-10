VERSION 5.00
Begin VB.Form eticket 
   Caption         =   "E-Ticket"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   11430
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Your Receipt Details"
      Height          =   2655
      Left            =   1440
      TabIndex        =   2
      Top             =   8160
      Width           =   7575
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   4200
         TabIndex        =   20
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Height          =   1575
         Left            =   600
         TabIndex        =   19
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label9 
         Caption         =   "Total Amount Payable"
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Ticket Charges"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Your Itinerary"
      Height          =   3975
      Left            =   1320
      TabIndex        =   1
      Top             =   3360
      Width           =   14055
      Begin VB.TextBox Text7 
         Height          =   1575
         Left            =   10680
         TabIndex        =   18
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox Text6 
         Height          =   1695
         Left            =   8160
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   1695
         Left            =   5280
         TabIndex        =   16
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox departing 
         Height          =   1695
         Left            =   2400
         TabIndex        =   15
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Flight Information"
         Height          =   255
         Left            =   10680
         TabIndex        =   11
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Status"
         Height          =   255
         Left            =   8160
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Arriving"
         Height          =   255
         Left            =   5280
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Departing"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Flight Number"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Passenger Ticket Information"
      Height          =   2175
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   14055
      Begin VB.TextBox Text10 
         Height          =   855
         Left            =   4080
         TabIndex        =   22
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   7080
         TabIndex        =   4
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox psgname 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label10 
         Caption         =   "Passenger Details"
         Height          =   255
         Left            =   4080
         TabIndex        =   21
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Ticket No."
         Height          =   255
         Left            =   7080
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Passenger Name"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Eticket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label11_Click()

End Sub

