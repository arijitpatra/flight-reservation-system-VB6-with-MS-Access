VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form eticket 
   BackColor       =   &H80000013&
   Caption         =   "E-Ticket"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton save 
      BackColor       =   &H0080FFFF&
      Caption         =   "SAVE && PRINT E-TICKET"
      BeginProperty Font 
         Name            =   "Nokia Standard Multiscript"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9600
      Width           =   6615
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H000000FF&
      Caption         =   "Rules"
      ForeColor       =   &H80000014&
      Height          =   4095
      Left            =   6600
      TabIndex        =   24
      Top             =   4920
      Width           =   9735
      Begin VB.Label Label10 
         BackColor       =   &H000000FF&
         Caption         =   $"eticket.frx":0000
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
         Height          =   2655
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   9255
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Your Receipt Details"
      Height          =   4095
      Left            =   2280
      TabIndex        =   2
      Top             =   4920
      Width           =   4215
      Begin VB.TextBox total 
         DataField       =   "Total Amount Payable"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   405
         Left            =   600
         TabIndex        =   19
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox ticketcharges 
         DataField       =   "Ticket Charges"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   1575
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "*all amounts are in INR"
         Height          =   255
         Left            =   2400
         TabIndex        =   33
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount Payable"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Ticket Charges"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Caption         =   "Your Itinerary"
      ForeColor       =   &H80000014&
      Height          =   2055
      Left            =   2280
      TabIndex        =   1
      Top             =   2760
      Width           =   14055
      Begin VB.TextBox status 
         Enabled         =   0   'False
         Height          =   405
         Left            =   8280
         TabIndex        =   32
         Text            =   "CONFIRMED"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox returnflightinfo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10560
         TabIndex        =   31
         Top             =   1320
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox returnstatus 
         Enabled         =   0   'False
         Height          =   375
         Left            =   8280
         TabIndex        =   30
         Text            =   "CONFIRMED"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox returnarrival 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5280
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox returndeparture 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   28
         Top             =   1320
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox returnflightnum 
         Enabled         =   0   'False
         Height          =   375
         Left            =   720
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox flightinfo 
         DataField       =   "Flight Information"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10560
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox arriving 
         DataField       =   "Arriving"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5280
         TabIndex        =   16
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox departing 
         DataField       =   "Departing"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   15
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox flightnum 
         DataField       =   "Flight Number"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Flight Information"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   10560
         TabIndex        =   11
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   8280
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Arriving"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   5280
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Departing"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Flight Number"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800080&
      Caption         =   "Passenger Ticket Information"
      ForeColor       =   &H80000014&
      Height          =   2175
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   14055
      Begin VB.TextBox dateandtime 
         DataField       =   "Booking Date and Time"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   22
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox bookingagent 
         DataField       =   "Booking Agent"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox pnr 
         DataField       =   "Ticket Number"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   405
         Left            =   480
         TabIndex        =   4
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox psgname 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   1605
         Left            =   3720
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   9735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Booking Date and Time"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Booking Agent"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ticket No."
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Passenger Name && Details"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   -1320
      Top             =   9720
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
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
      RecordSource    =   "etkt"
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
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   0
      Top             =   0
      Width           =   21000
   End
   Begin VB.Menu hme 
      Caption         =   "Home"
   End
   Begin VB.Menu fbk 
      Caption         =   "Feedback"
   End
   Begin VB.Menu ext 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Eticket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Eticket.Hide
Load welcome
welcome.Show
End Sub

Private Sub ext_Click()
Eticket.Hide
End Sub

Private Sub fbk_Click()
Eticket.Hide
Load feedback
feedback.Show
End Sub

Private Sub hme_Click()
Eticket.Hide
Load welcome
welcome.Show
End Sub

Private Sub save_Click()
'    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields(0) = bookingagent.Text
    Adodc1.Recordset.Fields(1) = dateandtime.Text
    Adodc1.Recordset.Fields(2) = pnr.Text
    Adodc1.Recordset.Fields(3) = psgname.Text
    Adodc1.Recordset.Fields(4) = flightnum.Text
    Adodc1.Recordset.Fields(5) = departing.Text
    Adodc1.Recordset.Fields(6) = arriving.Text
    Adodc1.Recordset.Fields(7) = status.Text
    Adodc1.Recordset.Fields(8) = flightinfo.Text
    Adodc1.Recordset.Fields(9) = ticketcharges.Text
    Adodc1.Recordset.Fields(10) = total.Text
    
    Adodc1.Recordset.Update
    Eticket.PrintForm
             
End Sub
