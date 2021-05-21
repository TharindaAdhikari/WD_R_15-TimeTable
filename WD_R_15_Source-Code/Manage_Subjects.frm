VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Manage_Subjects 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage_Subjects"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   14820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   15855
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Subjects"
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
         Left            =   1560
         TabIndex        =   21
         Top             =   360
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
      TabIndex        =   18
      Top             =   1200
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
         Left            =   360
         TabIndex        =   19
         Top             =   6960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Width           =   15855
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
         ItemData        =   "Manage_Subjects.frx":0000
         Left            =   8760
         List            =   "Manage_Subjects.frx":0002
         TabIndex        =   32
         Top             =   6240
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
         ItemData        =   "Manage_Subjects.frx":0004
         Left            =   8760
         List            =   "Manage_Subjects.frx":0006
         TabIndex        =   31
         Top             =   5520
         Width           =   3255
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
         ItemData        =   "Manage_Subjects.frx":0008
         Left            =   8760
         List            =   "Manage_Subjects.frx":000A
         TabIndex        =   30
         Top             =   4800
         Width           =   3255
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
         ItemData        =   "Manage_Subjects.frx":000C
         Left            =   8760
         List            =   "Manage_Subjects.frx":000E
         TabIndex        =   29
         Top             =   4200
         Width           =   3255
      End
      Begin VB.TextBox Text9 
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
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   24
         Top             =   6000
         WhatsThisHelpID =   1
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   5280
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   615
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
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   12
         Top             =   4800
         WhatsThisHelpID =   1
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
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
         Left            =   9240
         TabIndex        =   11
         Top             =   7080
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
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
         Left            =   10800
         TabIndex        =   10
         Top             =   7080
         Width           =   1335
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
         ItemData        =   "Manage_Subjects.frx":0010
         Left            =   2400
         List            =   "Manage_Subjects.frx":0012
         TabIndex        =   9
         Top             =   5400
         Width           =   3255
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Back"
         DisabledPicture =   "Manage_Subjects.frx":0014
         Height          =   495
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Manage_Subjects.frx":1A8F7
         Height          =   495
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   11895
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   2535
            Left            =   120
            TabIndex        =   0
            Top             =   120
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   4471
            _Version        =   393216
            BackColor       =   16777215
            BackColorFixed  =   16776960
            BackColorSel    =   65280
            BackColorBkg    =   -2147483635
            BackColorUnpopulated=   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "Serch Here........"
         Top             =   840
         Width           =   2295
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
         Left            =   2400
         TabIndex        =   4
         Top             =   4200
         WhatsThisHelpID =   1
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4080
         TabIndex        =   2
         Text            =   "Text6"
         Top             =   600
         Visible         =   0   'False
         Width           =   375
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
         Height          =   210
         Index           =   7
         Left            =   5760
         TabIndex        =   28
         Top             =   6240
         Width           =   2865
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
         Left            =   5520
         TabIndex        =   27
         Top             =   5520
         Width           =   2745
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
         Left            =   5640
         TabIndex        =   26
         Top             =   4800
         Width           =   2745
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
         Height          =   330
         Index           =   2
         Left            =   5640
         TabIndex        =   25
         Top             =   4200
         Width           =   2745
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
         Left            =   480
         TabIndex        =   17
         Top             =   6000
         Width           =   2205
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
         Left            =   600
         TabIndex        =   16
         Top             =   5400
         Width           =   1905
         WordWrap        =   -1  'True
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
         TabIndex        =   15
         Top             =   4800
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
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
         Left            =   240
         TabIndex        =   14
         Top             =   4200
         Width           =   2715
         WordWrap        =   -1  'True
      End
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
      Left            =   7800
      TabIndex        =   23
      Top             =   5760
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
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   2775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Manage_Subjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()


sql = "select * from Subject where ID='" + Text5.Text + "'"
Set dataset = mddata(sql)

With dataset

.Update
!Offerd_Year = Text3.Text
!Offered_Semester = Text8.Text
!Sub_name = Combo1.Text
!Sub_code = Text9.Text
!nu_of_lec_hours = Combo2.Text
!Nu_of_tut_hours = Combo3.Text
!Nu_of_lab_hours = Combo4.Text
!nu_of_Eva_hours = Combo5.Text
.Update

MsgBox "Successfully!", vbInformation

Unload Me
Me.Show
End With


End Sub

Private Sub Command3_Click()
If Not Text5.Text = "" Then
sql = "delete Subject where ID='" + Text5.Text + "'"
Set dataset = mddata(sql)
MsgBox "Done!!", vbInformation
Unload Me
Me.Show

Else
MsgBox "Select Subject!", vbInformation
MSHFlexGrid3.SetFocus
Exit Sub

End If
End Sub

Private Sub Command4_Click()
Unload Me
Home.Show
End Sub

Private Sub Command5_Click()
Unload Me
Subjects_Managment.Show
End Sub

Private Sub Form_Load()
sql = "select * from Subject "
Set dataset = mddata(sql)
With dataset

Set MSHFlexGrid3.DataSource = dataset


End With
End Sub

Private Sub MSHFlexGrid3_Click()
MSHFlexGrid3.Col = 1
Text5.Text = MSHFlexGrid3.Text
End Sub

Private Sub Text2_Change()


If Not Text2.Text = "" Then
sql = "select * from Subject where ID like '%" + Text2.Text + "%'"
Set dataset = mddata(sql)
With dataset

Set MSHFlexGrid3.DataSource = dataset

MSHFlexGrid3.ColWidth(0) = 0
MSHFlexGrid3.ColWidth(1) = 1400
MSHFlexGrid3.ColWidth(2) = 1000
MSHFlexGrid3.ColWidth(3) = 1800
MSHFlexGrid3.ColWidth(4) = 1800
MSHFlexGrid3.ColWidth(5) = 1200
MSHFlexGrid3.ColWidth(6) = 2000
MSHFlexGrid3.ColWidth(7) = 1800
MSHFlexGrid3.ColWidth(8) = 1000
MSHFlexGrid3.ColWidth(9) = 1600
End With


'End With
Else
sql = "select * from Subject "
Set dataset = mddata(sql)
With dataset

Set MSHFlexGrid3.DataSource = dataset

MSHFlexGrid3.ColWidth(0) = 0
MSHFlexGrid3.ColWidth(1) = 1400
MSHFlexGrid3.ColWidth(2) = 1000
MSHFlexGrid3.ColWidth(3) = 1800
MSHFlexGrid3.ColWidth(4) = 1800
MSHFlexGrid3.ColWidth(5) = 1200
MSHFlexGrid3.ColWidth(6) = 2000
MSHFlexGrid3.ColWidth(7) = 1800
MSHFlexGrid3.ColWidth(8) = 1000
MSHFlexGrid3.ColWidth(9) = 1600
End With
End If



End Sub

Private Sub Text2_GotFocus()
Text2.Text = ""
End Sub

Private Sub Text5_Change()

sql = "select * from Subject where ID = '" + Text5.Text + "'"
Set dataset = mddata(sql)
With dataset

If dataset.RecordCount > 0 Then
Text3.Text = !Offerd_Year
Text8.Text = !Offered_Semester
Combo1.Text = !Sub_name
Text9.Text = !Sub_code
Combo2.Text = !nu_of_lec_hours
Combo3.Text = !Nu_of_tut_hours
Combo4.Text = !Nu_of_lab_hours
Combo5.Text = !nu_of_Eva_hours


End If
End With
End Sub
