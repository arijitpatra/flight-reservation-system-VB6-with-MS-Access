VERSION 5.00
Begin VB.Form searchresult 
   Caption         =   "Available Flights"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   11130
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame return 
      BackColor       =   &H80000014&
      Caption         =   "Search Results (Round Trip)"
      Height          =   4815
      Left            =   3840
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   10695
      Begin VB.CommandButton return5 
         BackColor       =   &H0080C0FF&
         Caption         =   $"searchresult.frx":0000
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3720
         Width           =   10215
      End
      Begin VB.CommandButton return4 
         BackColor       =   &H00FFFF80&
         Caption         =   $"searchresult.frx":0151
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2880
         Width           =   10215
      End
      Begin VB.CommandButton return3 
         BackColor       =   &H00C0FFC0&
         Caption         =   $"searchresult.frx":02A6
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2040
         Width           =   10215
      End
      Begin VB.CommandButton return2 
         BackColor       =   &H00FFC0FF&
         Caption         =   $"searchresult.frx":03FC
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1200
         Width           =   10215
      End
      Begin VB.CommandButton return1 
         BackColor       =   &H0080FFFF&
         Caption         =   $"searchresult.frx":0550
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   10215
      End
   End
   Begin VB.Frame flights 
      BackColor       =   &H80000014&
      Caption         =   "Search Results (One Way)"
      Height          =   2895
      Left            =   3840
      TabIndex        =   0
      Top             =   960
      Width           =   10695
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   4200
         Width           =   10215
      End
      Begin VB.CommandButton flightfive 
         BackColor       =   &H000000FF&
         Caption         =   "Oceanbird Airlines | OA 555 | Delhi-Chennai | Daily Non Stop | Departure: 09:00 PM | Arrival: 12:00 AM "
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   10215
      End
      Begin VB.CommandButton flightfour 
         BackColor       =   &H000000FF&
         Caption         =   "Udaan India | UI 444 | Mumbai-Delhi | Daily Non Stop | Departure: 12:00 PM | Arrival: 03:00 PM"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1800
         Width           =   10215
      End
      Begin VB.CommandButton flightthree 
         BackColor       =   &H000000FF&
         Caption         =   "Sonic Air | SA 333 | Kolkata-Chennai | Daily Non Stop | Departure: 05:00 PM | Arrival: 08:00 PM"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1320
         Width           =   10215
      End
      Begin VB.CommandButton flighttwo 
         BackColor       =   &H000000FF&
         Caption         =   "Fly Jetways | FJ 222 | Kolkata-Mumbai | Daily Non Stop | Departure: 03:00 PM | Arrival: 04:30 PM"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   10215
      End
      Begin VB.CommandButton flightone 
         BackColor       =   &H000000FF&
         Caption         =   "Hello Airlines | HA 111 | Kolkata-Delhi | Daily Non Stop | Departure: 06:00 AM | Arrival: 08:00 AM  "
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   10215
      End
   End
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   0
      Picture         =   "searchresult.frx":06AD
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
Attribute VB_Name = "searchresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BackCmd_Click()
searchresult.Hide
Load journeydetails
journeydetails.Show
End Sub

Private Sub buyticket_Click()
searchresult.Hide
Load passengerdetails
passengerdetails.Show
End Sub

Private Sub contact_Click()
searchresult.Hide
Load contacts
contacts.Show
End Sub

Private Sub exit_Click()
searchresult.Hide
End Sub

Private Sub faq_Click()
searchresult.Hide
Load faskq
faskq.Show
End Sub

Private Sub flightfive_Click()
searchresult.Hide
Load passengerdetails
passengerdetails.Show
passengerdetails.flightchoosen.Text = "Your Flight Details: " & "Oceanbird Airlines | OA 555 | Delhi-Chennai | Daily Non Stop | Departure: 09:00 PM | Arrival: 12:00 AM  | " & "Date of Departure: " & journeydetails.dateofjourney.Text
passengerdetails.fareperadult.Text = "5600"
passengerdetails.fareperchild.Text = "5100"
passengerdetails.fareperseniorcitizen.Text = "5200"
passengerdetails.farepersplcon.Text = "4900"
passengerdetails.tax.Text = "1200"
passengerdetails.totalfare.Text = (Val(passengerdetails.fareperadult.Text) * Val(journeydetails.adults.Text)) + (Val(passengerdetails.fareperchild.Text) * Val(journeydetails.child.Text)) + (Val(passengerdetails.fareperseniorcitizen.Text) * Val(journeydetails.seniorcitizen.Text)) + (Val(passengerdetails.farepersplcon.Text) * Val(journeydetails.specialconcession.Text)) + (Val(journeydetails.adults.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.child.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.seniorcitizen.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.specialconcession.Text) * Val(passengerdetails.tax.Text))
Eticket.flightnum.Text = "OA 555"
Eticket.departing.Text = "09:00 PM" & " " & journeydetails.dateofjourney.Text
Eticket.arriving.Text = "12:00 AM" & " " & journeydetails.dateofjourney.Text
Eticket.flightinfo.Text = "Delhi-Chennai | Daily Non Stop"
End Sub

Private Sub flightfour_Click()
searchresult.Hide
Load passengerdetails
passengerdetails.Show
passengerdetails.flightchoosen.Text = "Your Flight Details: " & "Udaan India | UI 444 | Mumbai-Delhi | Daily Non Stop | Departure: 12:00 PM | Arrival: 03:00 PM | " & "Date of Departure: " & journeydetails.dateofjourney.Text
passengerdetails.fareperadult.Text = "5600"
passengerdetails.fareperchild.Text = "5100"
passengerdetails.fareperseniorcitizen.Text = "5200"
passengerdetails.farepersplcon.Text = "4900"
passengerdetails.tax.Text = "1200"
passengerdetails.totalfare.Text = (Val(passengerdetails.fareperadult.Text) * Val(journeydetails.adults.Text)) + (Val(passengerdetails.fareperchild.Text) * Val(journeydetails.child.Text)) + (Val(passengerdetails.fareperseniorcitizen.Text) * Val(journeydetails.seniorcitizen.Text)) + (Val(passengerdetails.farepersplcon.Text) * Val(journeydetails.specialconcession.Text)) + (Val(journeydetails.adults.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.child.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.seniorcitizen.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.specialconcession.Text) * Val(passengerdetails.tax.Text))
Eticket.flightnum.Text = "UI 444"
Eticket.departing.Text = "12:00 PM" & " " & journeydetails.dateofjourney.Text
Eticket.arriving.Text = "03:00 PM" & " " & journeydetails.dateofjourney.Text
Eticket.flightinfo.Text = "Mumbai-Delhi | Daily Non Stop"
End Sub

Private Sub flightone_Click()
searchresult.Hide
Load passengerdetails
passengerdetails.Show
passengerdetails.flightchoosen.Text = "Your Flight Details: " & "Hello Airlines | HA 111 | Kolkata-Delhi | Daily Non Stop | Departure: 06:00 AM | Arrival: 08:00 AM | " & "Date of Departure: " & journeydetails.dateofjourney.Text
passengerdetails.fareperadult.Text = "2600"
passengerdetails.fareperchild.Text = "2100"
passengerdetails.fareperseniorcitizen.Text = "2200"
passengerdetails.farepersplcon.Text = "1900"
passengerdetails.tax.Text = "1200"
passengerdetails.totalfare.Text = (Val(passengerdetails.fareperadult.Text) * Val(journeydetails.adults.Text)) + (Val(passengerdetails.fareperchild.Text) * Val(journeydetails.child.Text)) + (Val(passengerdetails.fareperseniorcitizen.Text) * Val(journeydetails.seniorcitizen.Text)) + (Val(passengerdetails.farepersplcon.Text) * Val(journeydetails.specialconcession.Text)) + (Val(journeydetails.adults.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.child.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.seniorcitizen.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.specialconcession.Text) * Val(passengerdetails.tax.Text))
Eticket.flightnum.Text = "HA 111"
Eticket.departing.Text = "06:00 AM" & " " & journeydetails.dateofjourney.Text
Eticket.arriving.Text = "08:00 AM" & " " & journeydetails.dateofjourney.Text
Eticket.flightinfo.Text = "Kolkata-Delhi | Daily Non Stop"
End Sub

Private Sub flightthree_Click()
searchresult.Hide
Load passengerdetails
passengerdetails.Show
passengerdetails.flightchoosen.Text = "Your Flight Details: " & "Sonic Air | SA 333 | Kolkata-Chennai | Daily Non Stop | Departure: 05:00 PM | Arrival: 08:00 PM | " & "Date of Departure: " & journeydetails.dateofjourney.Text
passengerdetails.fareperadult.Text = "3600"
passengerdetails.fareperchild.Text = "3100"
passengerdetails.fareperseniorcitizen.Text = "3200"
passengerdetails.farepersplcon.Text = "2900"
passengerdetails.tax.Text = "1200"
passengerdetails.totalfare.Text = (Val(passengerdetails.fareperadult.Text) * Val(journeydetails.adults.Text)) + (Val(passengerdetails.fareperchild.Text) * Val(journeydetails.child.Text)) + (Val(passengerdetails.fareperseniorcitizen.Text) * Val(journeydetails.seniorcitizen.Text)) + (Val(passengerdetails.farepersplcon.Text) * Val(journeydetails.specialconcession.Text)) + (Val(journeydetails.adults.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.child.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.seniorcitizen.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.specialconcession.Text) * Val(passengerdetails.tax.Text))
Eticket.flightnum.Text = "SA 333"
Eticket.departing.Text = "05:00 PM" & " " & journeydetails.dateofjourney.Text
Eticket.arriving.Text = "08:00 PM" & " " & journeydetails.dateofjourney.Text
Eticket.flightinfo.Text = "Kolkata-Chennai | Daily Non Stop"
End Sub

Private Sub flighttwo_Click()
searchresult.Hide
Load passengerdetails
passengerdetails.Show
passengerdetails.flightchoosen.Text = "Your Flight Details: " & "Fly Jetways | FJ 222 | Kolkata-Mumbai | Daily Non Stop | Departure: 03:00 PM | Arrival: 04:30 PM | " & "Date of Departure: " & journeydetails.dateofjourney.Text
passengerdetails.fareperadult.Text = "4600"
passengerdetails.fareperchild.Text = "4100"
passengerdetails.fareperseniorcitizen.Text = "4200"
passengerdetails.farepersplcon.Text = "3900"
passengerdetails.tax.Text = "1200"
passengerdetails.totalfare.Text = (Val(passengerdetails.fareperadult.Text) * Val(journeydetails.adults.Text)) + (Val(passengerdetails.fareperchild.Text) * Val(journeydetails.child.Text)) + (Val(passengerdetails.fareperseniorcitizen.Text) * Val(journeydetails.seniorcitizen.Text)) + (Val(passengerdetails.farepersplcon.Text) * Val(journeydetails.specialconcession.Text)) + (Val(journeydetails.adults.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.child.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.seniorcitizen.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.specialconcession.Text) * Val(passengerdetails.tax.Text))
Eticket.flightnum.Text = "FJ 222"
Eticket.departing.Text = "03:00 PM" & " " & journeydetails.dateofjourney.Text
Eticket.arriving.Text = "04:30 PM" & " " & journeydetails.dateofjourney.Text
Eticket.flightinfo.Text = "Kolkata-Mumbai | Daily Non Stop"
End Sub

Private Sub Form_Load()
If journeydetails.from.Text = "Kolkata" Then
If journeydetails.to.Text = "Delhi" Then
flightone.BackColor = &HC000&
End If
End If
If journeydetails.from.Text = "Kolkata" Then
If journeydetails.to.Text = "Mumbai" Then
flighttwo.BackColor = &HC000&
End If
End If
If journeydetails.from.Text = "Kolkata" Then
If journeydetails.to.Text = "Chennai" Then
flightthree.BackColor = &HC000&
End If
End If
If journeydetails.from.Text = "Mumbai" Then
If journeydetails.to.Text = "Delhi" Then
flightfour.BackColor = &HC000&
End If
End If
If journeydetails.from.Text = "Delhi" Then
If journeydetails.to.Text = "Chennai" Then
flightfive.BackColor = &HC000&
End If
End If
End Sub

Private Sub home_Click()
searchresult.Hide
Load welcome
welcome.Show
End Sub



Private Sub return1_Click()
searchresult.Hide
Load passengerdetails
passengerdetails.Show
passengerdetails.flightchoosen.Text = "Your Flight Details: " & "Hello Airlines | HA 111 | Kolkata-Delhi | Daily Non Stop | Departure: 06:00 AM | Arrival: 08:00 AM || Hello Airlines | HA 112 | Delhi-Kolkata | Daily Non Stop | Departure: 08:00 PM | Arrival: 10:00 PM  | " & " Date of Departure:  " & journeydetails.dateofjourney.Text & " | Date of Return: " & journeydetails.dateofreturn.Text
passengerdetails.fareperadult.Text = "5600"
passengerdetails.fareperchild.Text = "5100"
passengerdetails.fareperseniorcitizen.Text = "5200"
passengerdetails.farepersplcon.Text = "4900"
passengerdetails.tax.Text = "2200"
passengerdetails.totalfare.Text = (Val(passengerdetails.fareperadult.Text) * Val(journeydetails.adults.Text)) + (Val(passengerdetails.fareperchild.Text) * Val(journeydetails.child.Text)) + (Val(passengerdetails.fareperseniorcitizen.Text) * Val(journeydetails.seniorcitizen.Text)) + (Val(passengerdetails.farepersplcon.Text) * Val(journeydetails.specialconcession.Text)) + (Val(journeydetails.adults.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.child.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.seniorcitizen.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.specialconcession.Text) * Val(passengerdetails.tax.Text))
Eticket.flightnum.Text = "HA 111"
Eticket.returnflightnum.Text = "HA 112"
Eticket.departing.Text = "06:00 AM" & " " & journeydetails.dateofjourney.Text
Eticket.returndeparture.Text = "08:00 PM" & " " & journeydetails.dateofreturn.Text
Eticket.arriving.Text = "08:00 AM" & " " & journeydetails.dateofjourney.Text
Eticket.returnarrival.Text = "10:00 PM" & " " & journeydetails.dateofreturn.Text
Eticket.flightinfo.Text = "Kolkata-Delhi | Daily Non Stop"
Eticket.returnflightinfo.Text = "Delhi-Kolkata | Daily Non Stop"
End Sub

Private Sub return2_Click()
searchresult.Hide
Load passengerdetails
passengerdetails.Show
passengerdetails.flightchoosen.Text = "Your Flight Details: " & "Fly Jetways | FJ 222 | Kolkata-Mumbai | Daily Non Stop | Departure: 03:00 PM | Arrival: 05:00 PM || Fly Jetways | FJ 223 | Mumbai-Kolkata | Daily Non Stop | Departure: 07:00 PM | Arrival: 09:00 PM | " & " Date of Departure:  " & journeydetails.dateofjourney.Text & " | Date of Return: " & journeydetails.dateofreturn.Text
passengerdetails.fareperadult.Text = "9600"
passengerdetails.fareperchild.Text = "9100"
passengerdetails.fareperseniorcitizen.Text = "9200"
passengerdetails.farepersplcon.Text = "8900"
passengerdetails.tax.Text = "2200"
passengerdetails.totalfare.Text = (Val(passengerdetails.fareperadult.Text) * Val(journeydetails.adults.Text)) + (Val(passengerdetails.fareperchild.Text) * Val(journeydetails.child.Text)) + (Val(passengerdetails.fareperseniorcitizen.Text) * Val(journeydetails.seniorcitizen.Text)) + (Val(passengerdetails.farepersplcon.Text) * Val(journeydetails.specialconcession.Text)) + (Val(journeydetails.adults.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.child.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.seniorcitizen.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.specialconcession.Text) * Val(passengerdetails.tax.Text))
Eticket.flightnum.Text = "FJ 222"
Eticket.returnflightnum.Text = "FJ 223"
Eticket.departing.Text = "03:00 PM" & " " & journeydetails.dateofjourney.Text
Eticket.returndeparture.Text = "07:00 PM" & " " & journeydetails.dateofreturn.Text
Eticket.arriving.Text = "05:00 AM" & " " & journeydetails.dateofjourney.Text
Eticket.returnarrival.Text = "09:00 PM" & " " & journeydetails.dateofreturn.Text
Eticket.flightinfo.Text = "Kolkata-Mumbai | Daily Non Stop"
Eticket.returnflightinfo.Text = "Mumbai-Kolkata | Daily Non Stop"

End Sub

Private Sub return3_Click()
searchresult.Hide
Load passengerdetails
passengerdetails.Show
passengerdetails.flightchoosen.Text = "Your Flight Details: " & "Sonic Air | SA 333 | Kolkata-Chennai | Daily Non Stop | Departure: 05:00 PM | Arrival: 08:00 PM || Sonic Air | SA 332 | Chennai-Kolkata | Daily Non Stop | Departure: 08:00 AM | Arrival: 11:00 AM | " & " Date of Departure:  " & journeydetails.dateofjourney.Text & " | Date of Return: " & journeydetails.dateofreturn.Text
passengerdetails.fareperadult.Text = "7600"
passengerdetails.fareperchild.Text = "7100"
passengerdetails.fareperseniorcitizen.Text = "7200"
passengerdetails.farepersplcon.Text = "6900"
passengerdetails.tax.Text = "2200"
passengerdetails.totalfare.Text = (Val(passengerdetails.fareperadult.Text) * Val(journeydetails.adults.Text)) + (Val(passengerdetails.fareperchild.Text) * Val(journeydetails.child.Text)) + (Val(passengerdetails.fareperseniorcitizen.Text) * Val(journeydetails.seniorcitizen.Text)) + (Val(passengerdetails.farepersplcon.Text) * Val(journeydetails.specialconcession.Text)) + (Val(journeydetails.adults.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.child.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.seniorcitizen.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.specialconcession.Text) * Val(passengerdetails.tax.Text))
Eticket.flightnum.Text = "SA 333"
Eticket.returnflightnum.Text = "SA 332"
Eticket.departing.Text = "05:00 PM" & " " & journeydetails.dateofjourney.Text
Eticket.returndeparture.Text = "08:00 AM" & " " & journeydetails.dateofreturn.Text
Eticket.arriving.Text = "08:00 AM" & " " & journeydetails.dateofjourney.Text
Eticket.returnarrival.Text = "11:00 AM" & " " & journeydetails.dateofreturn.Text
Eticket.flightinfo.Text = "Kolkata-Chennai | Daily Non Stop"
Eticket.returnflightinfo.Text = "Chennai-Kolkata | Daily Non Stop"
End Sub

Private Sub return4_Click()
searchresult.Hide
Load passengerdetails
passengerdetails.Show
passengerdetails.flightchoosen.Text = "Your Flight Details: " & "Udaan India | UI 444 | Mumbai-Delhi | Daily Non Stop | Departure: 12:00 PM | Arrival: 03:00 PM || Udaan India | UI 442 | Delhi-Mumbai | Daily Non Stop | Departure: 01:00 PM | Arrival: 04:00 PM | " & " Date of Departure:  " & journeydetails.dateofjourney.Text & " | Date of Return: " & journeydetails.dateofreturn.Text
passengerdetails.fareperadult.Text = "11600"
passengerdetails.fareperchild.Text = "11100"
passengerdetails.fareperseniorcitizen.Text = "11200"
passengerdetails.farepersplcon.Text = "10900"
passengerdetails.tax.Text = "2200"
passengerdetails.totalfare.Text = (Val(passengerdetails.fareperadult.Text) * Val(journeydetails.adults.Text)) + (Val(passengerdetails.fareperchild.Text) * Val(journeydetails.child.Text)) + (Val(passengerdetails.fareperseniorcitizen.Text) * Val(journeydetails.seniorcitizen.Text)) + (Val(passengerdetails.farepersplcon.Text) * Val(journeydetails.specialconcession.Text)) + (Val(journeydetails.adults.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.child.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.seniorcitizen.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.specialconcession.Text) * Val(passengerdetails.tax.Text))
Eticket.flightnum.Text = "UI 444"
Eticket.returnflightnum.Text = "UI 442"
Eticket.departing.Text = "12:00 PM" & " " & journeydetails.dateofjourney.Text
Eticket.returndeparture.Text = "01:00 PM" & " " & journeydetails.dateofreturn.Text
Eticket.arriving.Text = "03:00 PM" & " " & journeydetails.dateofjourney.Text
Eticket.returnarrival.Text = "04:00 PM" & " " & journeydetails.dateofreturn.Text
Eticket.flightinfo.Text = "Mumbai-Delhi | Daily Non Stop"
Eticket.returnflightinfo.Text = "Delhi-Mumbai | Daily Non Stop"
End Sub

Private Sub return5_Click()
searchresult.Hide
Load passengerdetails
passengerdetails.Show
passengerdetails.flightchoosen.Text = "Your Flight Details: " & "Oceanbird Airlines | OA 555 | Delhi-Chennai | Daily Non Stop | Departure: 09:00 PM | Arrival: 12:00 AM || Oceanbird Airlines | OA 552 | Chennai-Delhi | Daily Non Stop | Departure: 10:00 AM | Arrival: 01:00 PM | " & " Date of Departure:  " & journeydetails.dateofjourney.Text & " | Date of Return: " & journeydetails.dateofreturn.Text
passengerdetails.fareperadult.Text = "10600"
passengerdetails.fareperchild.Text = "10100"
passengerdetails.fareperseniorcitizen.Text = "10200"
passengerdetails.farepersplcon.Text = "9900"
passengerdetails.tax.Text = "2200"
passengerdetails.totalfare.Text = (Val(passengerdetails.fareperadult.Text) * Val(journeydetails.adults.Text)) + (Val(passengerdetails.fareperchild.Text) * Val(journeydetails.child.Text)) + (Val(passengerdetails.fareperseniorcitizen.Text) * Val(journeydetails.seniorcitizen.Text)) + (Val(passengerdetails.farepersplcon.Text) * Val(journeydetails.specialconcession.Text)) + (Val(journeydetails.adults.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.child.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.seniorcitizen.Text) * Val(passengerdetails.tax.Text) + Val(journeydetails.specialconcession.Text) * Val(passengerdetails.tax.Text))
Eticket.flightnum.Text = "OA 555"
Eticket.returnflightnum.Text = "OA 552"
Eticket.departing.Text = "09:00 AM" & " " & journeydetails.dateofjourney.Text
Eticket.returndeparture.Text = "10:00 AM" & " " & journeydetails.dateofreturn.Text
Eticket.arriving.Text = "12:00 AM" & " " & journeydetails.dateofjourney.Text
Eticket.returnarrival.Text = "01:00 PM" & " " & journeydetails.dateofreturn.Text
Eticket.flightinfo.Text = "Delhi-Chennai | Daily Non Stop"
Eticket.returnflightinfo.Text = "Chennai-Delhi | Daily Non Stop"
End Sub

Private Sub sgno_Click()
searchresult.Hide
Load welcome
welcome.Show
End Sub

Private Sub signup_Click()
searchresult.Hide
Load sign
sign.Show
End Sub

Private Sub tandc_Click()
searchresult.Hide
Load tac
tac.Show
End Sub
