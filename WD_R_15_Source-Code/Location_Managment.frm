VERSION 5.00
Begin VB.Form Location_Managment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Location Managment"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   13815
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   27.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   12600
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to the Location Managment"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   0
         TabIndex        =   11
         Top             =   360
         Width           =   11775
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   -480
      TabIndex        =   9
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   9975
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "Location_Managment.frx":0000
         Height          =   1695
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Height          =   1695
         Left            =   5160
         TabIndex        =   3
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Location_Managment.frx":1A8E3
         Height          =   495
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Height          =   1695
         Left            =   3360
         TabIndex        =   1
         Top             =   4440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add Location"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         TabIndex        =   8
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Manage Location"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5160
         TabIndex        =   7
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "View Student Group"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3720
         TabIndex        =   6
         Top             =   6240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "WD_R_15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8280
         TabIndex        =   5
         Top             =   6240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Location_Managment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Add_Location.Show
End Sub

Private Sub Command2_Click()
Unload Me
Manage_Location.Show
End Sub

Private Sub Command4_Click()
Unload Me
Home.Show
End Sub
