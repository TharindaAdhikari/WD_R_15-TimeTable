VERSION 5.00

Begin VB.Form Add_New_Subject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add_New_Subject"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo6 
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
      ItemData        =   "Add_New_Subject.frx":0000
      Left            =   4920
      List            =   "Add_New_Subject.frx":0002
      TabIndex        =   29
      Top             =   7560
      Width           =   3255
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   14535
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add New Subject"
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
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   12255
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   2280
      TabIndex        =   4
      Top             =   1320
      Width           =   15855
      Begin VB.TextBox Text2 
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
         TabIndex        =   28
         Top             =   3480
         WhatsThisHelpID =   1
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Add_New_Subject.frx":0004
         Height          =   495
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Back"
         DisabledPicture =   "Add_New_Subject.frx":1A8E7
         Height          =   495
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "Add_New_Subject.frx":351CA
         Left            =   2640
         List            =   "Add_New_Subject.frx":351CC
         TabIndex        =   12
         Top             =   2760
         Width           =   3255
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
         Left            =   7200
         TabIndex        =   11
         Top             =   7200
         Width           =   1335
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
         Left            =   5520
         TabIndex        =   10
         Top             =   7200
         Width           =   1335
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
         TabIndex        =   9
         Top             =   2160
         WhatsThisHelpID =   1
         Width           =   2535
      End
      Begin VB.ComboBox Combo3 
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
         ItemData        =   "Add_New_Subject.frx":351CE
         Left            =   2640
         List            =   "Add_New_Subject.frx":351D0
         TabIndex        =   8
         Top             =   4200
         Width           =   3255
      End
      Begin VB.ComboBox Combo4 
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
         ItemData        =   "Add_New_Subject.frx":351D2
         Left            =   2640
         List            =   "Add_New_Subject.frx":351D4
         TabIndex        =   7
         Top             =   4920
         Width           =   3255
      End
      Begin VB.ComboBox Combo5 
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
         ItemData        =   "Add_New_Subject.frx":351D6
         Left            =   2640
         List            =   "Add_New_Subject.frx":351D8
         TabIndex        =   6
         Top             =   5640
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   7080
         TabIndex        =   5
         Top             =   4440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label O 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Offerd Year"
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
         Index           =   18
         Left            =   600
         TabIndex        =   23
         Top             =   1440
         Width           =   2385
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
         TabIndex        =   22
         Top             =   6480
         Width           =   1455
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Offered Semester"
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
         Left            =   480
         TabIndex        =   21
         Top             =   2160
         Width           =   2115
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Name"
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
         Left            =   570
         TabIndex        =   20
         Top             =   2880
         Width           =   2385
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Code"
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
         Index           =   1
         Left            =   600
         TabIndex        =   19
         Top             =   3480
         Width           =   2505
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Lecture Hours"
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
         Left            =   -105
         TabIndex        =   18
         Top             =   4200
         Width           =   3015
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number Of Tutorial Hours"
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
         Left            =   -240
         TabIndex        =   17
         Top             =   4920
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number Of Lab Hours"
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
         Left            =   0
         TabIndex        =   16
         Top             =   5640
         Width           =   2985
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number Of Evaluation Hours"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   7
         Left            =   360
         TabIndex        =   15
         Top             =   6240
         Width           =   2115
         WordWrap        =   -1  'True
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
      Top             =   1320
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
         Left            =   480
         TabIndex        =   3
         Top             =   7200
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
      ItemData        =   "Add_New_Subject.frx":351DA
      Left            =   5040
      List            =   "Add_New_Subject.frx":351DC
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
      TabIndex        =   27
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
      TabIndex        =   26
      Top             =   0
      Width           =   2775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Add_New_Subject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

If Text8.Text = "" Then
MsgBox "Enter the Offered Year!"
Text8.SetFocus
Exit Sub
End If

If Text1.Text = "" Then
MsgBox "Enter the Offered Semester!"
Text1.SetFocus
Exit Sub
End If

If Combo1.Text = "" Then
MsgBox "Enter the Subject Name!"
Combo1.SetFocus
Exit Sub
End If

If Text2.Text = "" Then
MsgBox "Enter the Subject Code!"
Text2.SetFocus
Exit Sub
End If

If Combo3.Text = "" Then
MsgBox "Enter the Number of Lecture Hours!"
Combo3.SetFocus
Exit Sub
End If

If Combo4.Text = "" Then
MsgBox "Enter the Number Of Tutorial Hours!"
Combo4.SetFocus
Exit Sub
End If

If Combo5.Text = "" Then
MsgBox "Enter the Number Of Lab Hours!"
Combo5.SetFocus
Exit Sub
End If

If Combo6.Text = "" Then
MsgBox "Enter the Number Of Evaluation Hours!"
Combo6.SetFocus
Exit Sub
End If




sql = "select * from Subject"
Set dataset = mddata(sql)
With dataset
.AddNew
!Offerd_Year = Text8.Text
!Offered_Semester = Text1.Text
!Sub_name = Combo1.Text
!Sub_code = Text2.Text
!nu_of_lec_hours = Combo3.Text
!Nu_of_tut_hours = Combo4.Text
!Nu_of_lab_hours = Combo5.Text
!nu_of_Eva_hours = Combo6.Text

.Update
MsgBox "Done!!", vbInformation
End With
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
Subjects_Managment.Show
End Sub
