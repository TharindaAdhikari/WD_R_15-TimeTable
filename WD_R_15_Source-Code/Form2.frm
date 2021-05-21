VERSION 5.00
Begin VB.Form Add_student 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Student Group "
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11325
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11325
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   -240
      TabIndex        =   3
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
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add Student Group "
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
         TabIndex        =   4
         Top             =   240
         Width           =   11775
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H008B8B00&
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   -960
      TabIndex        =   2
      Top             =   840
      Width           =   2535
      Begin VB.Label Label2 
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
         Left            =   960
         TabIndex        =   23
         Top             =   6600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   10695
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
         ItemData        =   "Form2.frx":0000
         Left            =   2760
         List            =   "Form2.frx":001C
         TabIndex        =   0
         Top             =   1320
         WhatsThisHelpID =   2
         Width           =   2535
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Back"
         DisabledPicture =   "Form2.frx":0050
         Height          =   495
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Form2.frx":1A933
         Height          =   495
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   480
         Width           =   1095
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
         Left            =   1680
         TabIndex        =   20
         Top             =   5400
         WhatsThisHelpID =   5
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
         Left            =   3600
         TabIndex        =   19
         Top             =   5400
         WhatsThisHelpID =   6
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Genarate ID"
         ForeColor       =   &H80000008&
         Height          =   5055
         Left            =   5760
         TabIndex        =   13
         Top             =   1200
         Width           =   3615
         Begin VB.CommandButton Command1 
            Caption         =   "Genarate ID"
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
            Left            =   1680
            TabIndex        =   18
            Top             =   3960
            WhatsThisHelpID =   9
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   1680
            TabIndex        =   17
            Top             =   1920
            WhatsThisHelpID =   8
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   1680
            TabIndex        =   16
            Top             =   960
            WhatsThisHelpID =   7
            Width           =   1815
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Group ID"
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
            Index           =   4
            Left            =   -120
            TabIndex        =   15
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Group ID"
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
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   1455
         End
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
         ItemData        =   "Form2.frx":35216
         Left            =   2760
         List            =   "Form2.frx":35229
         TabIndex        =   8
         Top             =   3000
         WhatsThisHelpID =   3
         Width           =   2535
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
         ItemData        =   "Form2.frx":35242
         Left            =   2760
         List            =   "Form2.frx":35255
         TabIndex        =   7
         Top             =   3840
         WhatsThisHelpID =   4
         Width           =   2535
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
         ItemData        =   "Form2.frx":35268
         Left            =   2760
         List            =   "Form2.frx":35278
         TabIndex        =   6
         Top             =   2160
         WhatsThisHelpID =   2
         Width           =   2535
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Academic Year Semester"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   2505
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Group Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   480
         TabIndex        =   12
         Top             =   3840
         Width           =   1815
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
         Index           =   1
         Left            =   840
         TabIndex        =   11
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Programme"
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
         Index           =   0
         Left            =   960
         TabIndex        =   10
         Top             =   2280
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Add_student"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo4.Text = "" Then
MsgBox "Enter the year Semester! ", vbInformation
Combo4.SetFocus
Exit Sub
End If

If Combo1.Text = "" Then
MsgBox "Enter the Programme! ", vbInformation
Combo1.SetFocus
Exit Sub
End If

If Combo2.Text = "" Then
MsgBox "Enter the Group Number! ", vbInformation
Combo2.SetFocus
Exit Sub
End If

Text1.Text = Combo4.Text + "." + Combo1.Text + "." + Combo3.Text
Text2.Text = Combo4.Text + "." + Combo1.Text + "." + Combo3.Text + "." + Combo2.Text



End Sub

Private Sub Command2_Click()

If Combo4.Text = "" Then
MsgBox "Enter the Year Semester!", vbInformation
Combo4.SetFocus
Exit Sub

End If

If Combo1.Text = "" Then
MsgBox "Enter the Programme!", vbInformation
Combo1.SetFocus
Exit Sub

End If

If Combo3.Text = "" Then
MsgBox "Enter the Group Number!", vbInformation
Combo3.SetFocus
Exit Sub

End If

If Combo2.Text = "" Then
MsgBox "Enter the Sub Group Number!", vbInformation
Combo2.SetFocus
Exit Sub

End If

If Text1.Text = "" Then
MsgBox "Genarate the group Id!", vbInformation
Command1.SetFocus
Exit Sub

End If

If Text2.Text = "" Then
MsgBox "Genarate the Sub Group Id!", vbInformation
Command1.SetFocus
Exit Sub

End If

sql = "select * from Student_group"
Set dataset = mddata(sql)
With dataset
.AddNew
!year_semester = Combo4.Text
!programm = Combo1.Text
!group_number = Combo3.Text
!sub_group_number = Combo2.Text
!group_id = Text1.Text
!sub_group_id = Text2.Text
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
Student_Group_Managment.Show
End Sub
