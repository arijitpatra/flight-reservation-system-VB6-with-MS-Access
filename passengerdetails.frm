VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form passengerdetails 
   Caption         =   "Passenger Details"
   ClientHeight    =   3180
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Accept T&&C and click on ""Proceed"""
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Enter correct details in the fields and click on ""Done"""
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Click on ""Enter Details"""
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Click on ""OK! Continue"""
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox flightchoosen 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   55
      Top             =   1800
      Width           =   13695
   End
   Begin VB.Frame fare 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculate Fare"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   480
      TabIndex        =   54
      Top             =   1680
      Width           =   4215
      Begin VB.CommandButton continue 
         BackColor       =   &H0000FFFF&
         Caption         =   "OK! Continue"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox totalfare 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   66
         Top             =   4560
         Width           =   3255
      End
      Begin VB.TextBox fareperchild 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   64
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox farepersplcon 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   62
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox fareperseniorcitizen 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   60
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox tax 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   57
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox fareperadult 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   56
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "*all fares and taxes in INR"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   73
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "YOUR TOTAL PAYABLE FARE INCLUDING        ALL TAXES && SURCHARGES IS INR:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   67
         Top             =   3840
         Width           =   3255
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fare/Child :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fare/Spl. Concession Holders :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   63
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fare/Sr. Citizen :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   61
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Taxes/    Person :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   59
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fare/Adult :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H0080FFFF&
      Caption         =   "ENTER DETAILS"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   4800
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Show e-Ticket"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nokia Standard Multiscript"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   17280
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   120
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1215
      Left            =   9480
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2143
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
      RecordSource    =   "psgdetails"
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
   Begin VB.Frame passenger 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Passenger Details"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   4920
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   13695
      Begin VB.CommandButton done 
         BackColor       =   &H0000FFFF&
         Caption         =   "Done"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12240
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   4560
         Width           =   735
      End
      Begin VB.ComboBox foodchoice5 
         Height          =   315
         Left            =   12120
         TabIndex        =   50
         Text            =   "Type"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox foodcheck5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check4"
         Height          =   255
         Left            =   11400
         TabIndex        =   49
         Top             =   3600
         Width           =   255
      End
      Begin VB.TextBox phonenum5 
         Height          =   285
         Left            =   8880
         TabIndex        =   48
         Top             =   3600
         Width           =   2295
      End
      Begin VB.ComboBox class5 
         Height          =   315
         Left            =   7320
         TabIndex        =   47
         Text            =   "Select Class"
         Top             =   3600
         Width           =   1335
      End
      Begin VB.ComboBox gender5 
         Height          =   315
         Left            =   6000
         TabIndex        =   46
         Text            =   "Gender"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox age5 
         Height          =   285
         Left            =   5160
         TabIndex        =   45
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox surname5 
         Height          =   285
         Left            =   3360
         TabIndex        =   44
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox name5 
         Height          =   285
         Left            =   1440
         TabIndex        =   43
         Top             =   3600
         Width           =   1695
      End
      Begin VB.ComboBox foodchoice4 
         Height          =   315
         Left            =   12120
         TabIndex        =   42
         Text            =   "Type"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox foodcheck4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check3"
         Height          =   255
         Left            =   11400
         TabIndex        =   41
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox phonenum4 
         Height          =   285
         Left            =   8880
         TabIndex        =   40
         Top             =   3000
         Width           =   2295
      End
      Begin VB.ComboBox class4 
         Height          =   315
         Left            =   7320
         TabIndex        =   39
         Text            =   "Select Class"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox gender4 
         Height          =   315
         Left            =   6000
         TabIndex        =   38
         Text            =   "Gender"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox age4 
         Height          =   285
         Left            =   5160
         TabIndex        =   37
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox surname4 
         Height          =   285
         Left            =   3360
         TabIndex        =   36
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox name4 
         Height          =   285
         Left            =   1440
         TabIndex        =   35
         Top             =   3000
         Width           =   1695
      End
      Begin VB.ComboBox foodchoice3 
         Height          =   315
         Left            =   12120
         TabIndex        =   34
         Text            =   "Type"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox foodcheck3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check2"
         Height          =   315
         Left            =   11400
         TabIndex        =   33
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox phonenum3 
         Height          =   285
         Left            =   8880
         TabIndex        =   32
         Top             =   2400
         Width           =   2295
      End
      Begin VB.ComboBox class3 
         Height          =   315
         Left            =   7320
         TabIndex        =   31
         Text            =   "Select Class"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.ComboBox gender3 
         Height          =   315
         Left            =   6000
         TabIndex        =   30
         Text            =   "Gender"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox age3 
         Height          =   285
         Left            =   5160
         TabIndex        =   29
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox surname3 
         Height          =   285
         Left            =   3360
         TabIndex        =   28
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox name3 
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   2400
         Width           =   1695
      End
      Begin VB.ComboBox foodchoice2 
         Height          =   315
         Left            =   12120
         TabIndex        =   26
         Text            =   "Type"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox foodcheck2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   315
         Left            =   11400
         TabIndex        =   25
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox phonenum2 
         Height          =   285
         Left            =   8880
         TabIndex        =   24
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox class2 
         Height          =   315
         Left            =   7320
         TabIndex        =   23
         Text            =   "Select Class"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox gender2 
         Height          =   315
         Left            =   6000
         TabIndex        =   22
         Text            =   "Gender"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox age2 
         Height          =   285
         Left            =   5160
         TabIndex        =   21
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox surname2 
         Height          =   285
         Left            =   3360
         TabIndex        =   20
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox name2 
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox foodchoice1 
         DataField       =   "food"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   12120
         TabIndex        =   17
         Text            =   "Type"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox foodcheck1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check2"
         Height          =   315
         Left            =   11400
         TabIndex        =   16
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox phonenum1 
         DataField       =   "mobile number"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   8880
         TabIndex        =   13
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox class1 
         DataField       =   "class"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   7320
         TabIndex        =   10
         Text            =   "Select Class"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Proceed 
         BackColor       =   &H0000FFFF&
         Caption         =   "Proceed"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox accept 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8400
         TabIndex        =   5
         Top             =   5160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox gender1 
         DataField       =   "gender"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   6000
         TabIndex        =   4
         Text            =   "Gender"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox age1 
         DataField       =   "age"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   5160
         TabIndex        =   3
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox surname1 
         DataField       =   "surname"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   3360
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox name1 
         DataField       =   "name"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "I have read all the T&&C and I accept them"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8760
         TabIndex        =   18
         Top             =   5160
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Food"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11400
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mobile No."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8880
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Surname"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Label navback 
      BackStyle       =   0  'Transparent
      Caption         =   "<< Navigate"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   17160
      TabIndex        =   75
      Top             =   9480
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   $"passengerdetails.frx":0000
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   74
      Top             =   480
      Width           =   12375
   End
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   0
      Picture         =   "passengerdetails.frx":00B3
      Top             =   0
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
   Begin VB.Menu sgno 
      Caption         =   "Sign Out"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "passengerdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnew_Click()
Adodc1.Recordset.AddNew
passenger.Visible = True
name1.SetFocus
cmdnew.Visible = False
Command3.BackColor = &HC000&
Command4.Visible = True
End Sub
Private Sub Command1_Click()
'Eticket.Adodc1.Recordset.AddNew
passengerdetails.Hide
Eticket.Show
End Sub

Private Sub contact_Click()
passengerdetails.Hide
Load contacts
contacts.Show
End Sub

Private Sub continue_Click()
cmdnew.Visible = True
Command2.BackColor = &HC000&
Command3.Visible = True
Eticket.total.Text = passengerdetails.totalfare.Text
Eticket.ticketcharges.Text = "BASE FARES IN INR PER PERSON: " + "Adults: " + passengerdetails.fareperadult.Text + " Child: " + passengerdetails.fareperchild.Text + " Senior Citizen: " + passengerdetails.fareperseniorcitizen.Text + " Special Concession: " + passengerdetails.farepersplcon.Text
End Sub

Private Sub done_Click()
If name1.Text = "" Then
MsgBox ("Fill all the details")
ElseIf age1 = "" Then
MsgBox ("Fill all the details")
ElseIf gender1 = "" Then
MsgBox ("Fill all the details")
ElseIf class1 = "" Then
MsgBox ("Fill all the details")
ElseIf phonenum1 = "" Then
MsgBox ("Fill all the details")
Else
Proceed.Visible = True
Label8.Visible = True
accept.Visible = True
Command4.BackColor = &HC000&
Command5.Visible = True
End If
End Sub

Private Sub exit_Click()
passengerdetails.Hide
End Sub

Private Sub faq_Click()
passengerdetails.Hide
Load faskq
faskq.Show
End Sub

Private Sub foodcheck1_Click()
If foodcheck1.Value = 1 Then
foodchoice1.Visible = True
Else
foodchoice1.Visible = False

End If
End Sub

Private Sub foodcheck2_Click()
If foodcheck2.Value = 1 Then
foodchoice2.Visible = True
Else
foodchoice2.Visible = False

End If
End Sub

Private Sub foodcheck3_Click()
If foodcheck3.Value = 1 Then
foodchoice3.Visible = True
Else
foodchoice3.Visible = False

End If
End Sub

Private Sub foodcheck4_Click()
If foodcheck4.Value = 1 Then
foodchoice4.Visible = True
Else
foodchoice4.Visible = False

End If
End Sub

Private Sub foodcheck5_Click()
If foodcheck5.Value = 1 Then
foodchoice5.Visible = True
Else
foodchoice5.Visible = False

End If
End Sub

Private Sub Form_Load()
With gender1
.AddItem ("Male")
.AddItem ("Female")
.AddItem ("Other")
End With

With gender2
.AddItem ("Male")
.AddItem ("Female")
.AddItem ("Other")
End With

With gender3
.AddItem ("Male")
.AddItem ("Female")
.AddItem ("Other")
End With

With gender4
.AddItem ("Male")
.AddItem ("Female")
.AddItem ("Other")
End With

With gender5
.AddItem ("Male")
.AddItem ("Female")
.AddItem ("Other")
End With

With foodchoice1
.AddItem ("Veg")
.AddItem ("Non Veg")
End With

With foodchoice2
.AddItem ("Veg")
.AddItem ("Non Veg")
End With

With foodchoice3
.AddItem ("Veg")
.AddItem ("Non Veg")
End With

With foodchoice4
.AddItem ("Veg")
.AddItem ("Non Veg")
End With

With foodchoice5
.AddItem ("Veg")
.AddItem ("Non Veg")
End With

With class1
.AddItem ("Business ")
.AddItem ("Economy ")
End With

With class2
.AddItem ("Business ")
.AddItem ("Economy ")
End With

With class3
.AddItem ("Business ")
.AddItem ("Economy ")
End With

With class4
.AddItem ("Business ")
.AddItem ("Economy ")
End With

With class5
.AddItem ("Business ")
.AddItem ("Economy ")
End With
End Sub

Private Sub home_Click()
passengerdetails.Hide
Load welcome
welcome.Show
End Sub




Private Sub navback_Click()
passengerdetails.Hide
Load searchresult
searchresult.Show
End Sub

Private Sub Proceed_Click()

If accept.Value = 1 Then
Adodc1.Recordset.Fields(0) = name1.Text
    Adodc1.Recordset.Fields(1) = surname1.Text
    Adodc1.Recordset.Fields(2) = age1.Text
    Adodc1.Recordset.Fields(3) = gender1.Text
    Adodc1.Recordset.Fields(4) = class1.Text
    Adodc1.Recordset.Fields(5) = phonenum1.Text
    Adodc1.Recordset.Fields(6) = foodchoice1.Text
    Adodc1.Recordset.Update

fare.Visible = False
Command5.BackColor = &HC000&
Eticket.psgname.Text = "Passenger 1: " + passengerdetails.name1.Text + " " + passengerdetails.surname1.Text + " " + passengerdetails.age1.Text + " " + passengerdetails.gender1.Text + " " + passengerdetails.class1.Text + " " + passengerdetails.phonenum1.Text + " " + passengerdetails.foodchoice1.Text + " Passenger 2: " + passengerdetails.name2.Text + " " + passengerdetails.surname2.Text + " " + passengerdetails.age2.Text + " " + passengerdetails.gender2.Text + " " + passengerdetails.class2.Text + " " + passengerdetails.phonenum2.Text + " " + passengerdetails.foodchoice2.Text + " Passenger 3: " + passengerdetails.name3.Text + " " + passengerdetails.surname3.Text + " " + passengerdetails.age3.Text + " " + passengerdetails.gender3.Text + " " + passengerdetails.class3.Text + " " + passengerdetails.phonenum3.Text + " " + passengerdetails.foodchoice3.Text
Eticket.dateandtime.Text = Format(Now, "dddd,mmmm,dd,yyyy hh:mm:ss")

passenger.Visible = False
flightchoosen.Visible = False
MsgBox ("Thank You!")
  Command2.Visible = False
  Command3.Visible = False
  Command4.Visible = False
  Command5.Visible = False
  
Eticket.pnr.Text = "ABC1234567890"

  
  'Load Eticket
  'Eticket.Show
  Command1.BackColor = &HC000&
  Command1.Enabled = True
  
  
Else
MsgBox ("Accept T&C to proceed!")
End If
         
  End Sub

Private Sub sgno_Click()
passengerdetails.Hide
Load welcome
welcome.Show
End Sub

Private Sub signup_Click()
passengerdetails.Hide
Load sign
sign.Show
End Sub

Private Sub tandc_Click()
passengerdetails.Hide
Load tac
tac.Show
End Sub


