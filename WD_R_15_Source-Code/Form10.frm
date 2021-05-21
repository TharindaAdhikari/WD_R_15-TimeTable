VERSION 5.00
Begin VB.Form Add_Location 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add_Location"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11040
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   2400
      TabIndex        =   8
      Top             =   1320
      Width           =   15855
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   195
         Left            =   6000
         TabIndex        =   25
         Top             =   3120
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   3600
         TabIndex        =   24
         Top             =   3120
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   7320
         TabIndex        =   15
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2760
         TabIndex        =   14
         Top             =   3600
         WhatsThisHelpID =   1
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   13
         Top             =   2160
         WhatsThisHelpID =   1
         Width           =   2535
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2640
         TabIndex        =   0
         Top             =   1440
         WhatsThisHelpID =   1
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         TabIndex        =   12
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6600
         TabIndex        =   11
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Back"
         DisabledPicture =   "Form10.frx":0000
         Height          =   495
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Form10.frx":1A8E3
         Height          =   495
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratory"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   5880
         TabIndex        =   23
         Top             =   3120
         Width           =   1995
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lecture Hall"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   2910
         TabIndex        =   22
         Top             =   3120
         Width           =   3195
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   360
         TabIndex        =   21
         Top             =   3720
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   360
         TabIndex        =   20
         Top             =   3120
         Width           =   2655
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   600
         TabIndex        =   19
         Top             =   2160
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulding Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   600
         TabIndex        =   18
         Top             =   1440
         Width           =   1995
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
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
         Left            =   11640
         TabIndex        =   17
         Top             =   6480
         Width           =   1455
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Room Bulding Wise"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   18
         Left            =   585
         TabIndex        =   16
         Top             =   480
         Width           =   3195
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14535
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add Location"
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
         TabIndex        =   5
         Top             =   240
         Width           =   12255
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   2535
      Begin VB.Label Label1 
         BackColor       =   &H008B8B00&
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
         Left            =   600
         TabIndex        =   3
         Top             =   6600
         Width           =   1455
      End
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Form10.frx":351C6
      Left            =   5040
      List            =   "Form10.frx":351C8
      TabIndex        =   1
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label Label152 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Group Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label152 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   0
      Width           =   2775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Add_Location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_LostFocus()
If Check1.Value = 1 Then
Check2.Enabled = False

Else
Check2.Enabled = True
End If
End Sub

Private Sub Check2_LostFocus()
If Check2.Value = 1 Then
Check1.Enabled = False

Else
Check1.Enabled = True
End If
End Sub

Private Sub Command2_Click()

If Text8.Text = "" Then

MsgBox "Enter the Buliding Name!", vbInformation
Text8.SetFocus
Exit Sub
End If


If Text1.Text = "" Then

MsgBox "Enter the Room Name!", vbInformation
Text1.SetFocus
Exit Sub
End If

If Text3.Text = "" Then

MsgBox "Enter the Capacity!", vbInformation
Text3.SetFocus
Exit Sub
End If

If Check1.Value = 0 And Check2.Value = 0 Then

MsgBox "Enter the Room Type!", vbInformation
Check1.SetFocus
Exit Sub
End If


sql = "select * from Loacation"
Set dataset = mddata(sql)

With dataset

.AddNew
!Bulding_Name = Text8.Text
!Room_Name = Text1.Text

If Check1.Value = 1 Then
!Room_type = "Lecture Hall"
End If

If Check2.Value = 1 Then
!Room_type = "Laboratory"
End If

!capacity = Text3.Text

.Update
End With
MsgBox "Successfully!", vbInformation

Unload Me
Me.Show
End Sub

Private Sub Command3_Click()
Unload Me
Me.Show
End Sub

Private Sub Command4_Click()
Unload Me
Home.Show
End Sub

Private Sub Command5_Click()
Unload Me
Location_Managment.Show
End Sub
