VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form manage_student_group 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Student Group"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14820
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   14820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   2280
      TabIndex        =   7
      Top             =   1080
      Width           =   15855
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
         Left            =   7320
         TabIndex        =   30
         Top             =   7080
         WhatsThisHelpID =   9
         Width           =   1695
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
         ItemData        =   "Form7.frx":0000
         Left            =   2400
         List            =   "Form7.frx":0010
         TabIndex        =   29
         Top             =   4800
         WhatsThisHelpID =   2
         Width           =   2535
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
         ItemData        =   "Form7.frx":0027
         Left            =   2400
         List            =   "Form7.frx":0043
         TabIndex        =   28
         Top             =   4080
         WhatsThisHelpID =   2
         Width           =   2535
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Genarate ID"
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   7680
         TabIndex        =   23
         Top             =   4320
         Width           =   4335
         Begin VB.TextBox Text7 
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
            Left            =   1920
            TabIndex        =   25
            Top             =   960
            WhatsThisHelpID =   7
            Width           =   1935
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
            Left            =   1920
            TabIndex        =   24
            Top             =   1920
            WhatsThisHelpID =   8
            Width           =   1935
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
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   1455
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
            Left            =   120
            TabIndex        =   26
            Top             =   1920
            Width           =   1455
         End
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4080
         TabIndex        =   18
         Text            =   "Text6"
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
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
         ItemData        =   "Form7.frx":0077
         Left            =   2400
         List            =   "Form7.frx":0079
         TabIndex        =   16
         Top             =   6000
         Width           =   3255
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
         TabIndex        =   15
         Text            =   "Serch Here........"
         Top             =   840
         Width           =   2295
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   240
         TabIndex        =   14
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
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Form7.frx":007B
         Height          =   495
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Back"
         DisabledPicture =   "Form7.frx":1A95E
         Height          =   495
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   12
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
         ItemData        =   "Form7.frx":35241
         Left            =   2400
         List            =   "Form7.frx":35243
         TabIndex        =   11
         Top             =   5400
         Width           =   3255
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
         TabIndex        =   9
         Top             =   7080
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   5280
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acedemic Year"
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
         Left            =   -75
         TabIndex        =   22
         Top             =   4200
         Width           =   2685
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Height          =   210
         Index           =   10
         Left            =   705
         TabIndex        =   21
         Top             =   4800
         Width           =   1665
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Height          =   210
         Index           =   0
         Left            =   705
         TabIndex        =   20
         Top             =   5400
         Width           =   1875
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Height          =   210
         Index           =   1
         Left            =   345
         TabIndex        =   19
         Top             =   6000
         Width           =   2175
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
      TabIndex        =   3
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
         TabIndex        =   4
         Top             =   6960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15855
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Student Group"
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
         TabIndex        =   2
         Top             =   360
         Width           =   12255
      End
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
      TabIndex        =   6
      Top             =   0
      Width           =   2775
      WordWrap        =   -1  'True
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
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
End
Attribute VB_Name = "manage_student_group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Command2.Enabled = False
End Sub

Private Sub Combo1_LostFocus()
Command2.Enabled = False
End Sub

Private Sub Combo2_Change()
Command2.Enabled = False
End Sub

Private Sub Combo2_LostFocus()
Command2.Enabled = False
End Sub

Private Sub Combo3_Change()
Command2.Enabled = False
End Sub

Private Sub Combo3_LostFocus()
Command2.Enabled = False
End Sub

Private Sub Combo4_Change()
Command2.Enabled = False
End Sub

Private Sub Combo4_LostFocus()
Command2.Enabled = False
End Sub

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

If Combo3.Text = "" Then
MsgBox "Enter the Sub Group Number! ", vbInformation
Combo2.SetFocus
Exit Sub
End If

Text7.Text = Combo4.Text + "." + Combo2.Text + "." + Combo1.Text
Text1.Text = Combo4.Text + "." + Combo2.Text + "." + Combo1.Text + "." + Combo3.Text

Command2.Enabled = True


End Sub

Private Sub Command2_Click()
If Combo4.Text = "" Then
MsgBox "Enter the Acedemic Year", vbInformation
Combo4.SetFocus
Exit Sub
End If

If Combo2.Text = "" Then
MsgBox "Enter the Programme", vbInformation
Combo2.SetFocus
Exit Sub
End If


sql = "select * from Student_group where ID='" + Text5.Text + "'"
Set dataset = mddata(sql)

With dataset

.Update
!year_semester = Combo4.Text
!programm = Combo2.Text
!group_number = Combo1.Text
!sub_group_number = Combo3.Text
!group_id = Text7.Text
!sub_group_id = Text1.Text
.Update

MsgBox "Successfully!", vbInformation

Unload Me
Me.Show
End With



End Sub

Private Sub Command3_Click()
If Not Text5.Text = "" Then
sql = "delete Student_group where ID='" + Text5.Text + "'"
Set dataset = mddata(sql)
MsgBox "Done!!", vbInformation
Unload Me
Me.Show

Else
MsgBox "Select Student Group!", vbInformation
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
Student_Group_Managment.Show
End Sub

Private Sub Form_Load()
sql = "select * from Student_group "
Set dataset = mddata(sql)
With dataset

Set MSHFlexGrid3.DataSource = dataset
MSHFlexGrid3.ColWidth(0) = 0
MSHFlexGrid3.ColWidth(1) = 0
MSHFlexGrid3.ColWidth(2) = 1800
MSHFlexGrid3.ColWidth(3) = 1800
MSHFlexGrid3.ColWidth(4) = 1800
MSHFlexGrid3.ColWidth(5) = 2200
MSHFlexGrid3.ColWidth(6) = 2000
MSHFlexGrid3.ColWidth(7) = 1800
MSHFlexGrid3.ColWidth(8) = 1000
MSHFlexGrid3.ColWidth(9) = 1600

End With
End Sub

Private Sub MSHFlexGrid3_Click()
MSHFlexGrid3.Col = 1
Text5.Text = MSHFlexGrid3.Text
End Sub

Private Sub Text2_Change()

If Not Text2.Text = "" Then
sql = "select * from Student_group where ID like '%" + Text2.Text + "%'"
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
MSHFlexGrid3.ColWidth(7) = 0
MSHFlexGrid3.ColWidth(8) = 1000
MSHFlexGrid3.ColWidth(9) = 1600
End With


'End With
Else
sql = "select * from Student_group "
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
MSHFlexGrid3.ColWidth(7) = 0
MSHFlexGrid3.ColWidth(8) = 1000
MSHFlexGrid3.ColWidth(9) = 1600
End With
End If

End Sub

Private Sub Text2_GotFocus()
Text2.Text = ""
End Sub

Private Sub Text5_Change()
sql = "select * from Student_group where ID = '" + Text5.Text + "'"
Set dataset = mddata(sql)
With dataset

If dataset.RecordCount > 0 Then
Combo4.Text = !year_semester
Combo2.Text = !programm
Combo1.Text = !group_number
Combo3.Text = !sub_group_number
Text7.Text = !group_id
Text1.Text = !sub_group_id


End If
End With
End Sub
