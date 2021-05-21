VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form manage_lec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Lecturer"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14895
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   14895
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
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
         Caption         =   "Manage Lecturer"
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   15855
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4080
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1920
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
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
         ItemData        =   "Form6.frx":0000
         Left            =   8880
         List            =   "Form6.frx":0016
         TabIndex        =   30
         Top             =   5640
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
         ItemData        =   "Form6.frx":007E
         Left            =   8880
         List            =   "Form6.frx":0080
         TabIndex        =   29
         Top             =   5040
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
         ItemData        =   "Form6.frx":0082
         Left            =   8880
         List            =   "Form6.frx":0098
         TabIndex        =   28
         Top             =   4320
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
         ItemData        =   "Form6.frx":00CE
         Left            =   2400
         List            =   "Form6.frx":00D0
         TabIndex        =   27
         Top             =   6000
         Width           =   3255
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
         TabIndex        =   26
         Top             =   4200
         WhatsThisHelpID =   1
         Width           =   2535
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
         TabIndex        =   25
         Text            =   "Serch Here........"
         Top             =   840
         Width           =   2295
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   240
         TabIndex        =   24
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
         DisabledPicture =   "Form6.frx":00D2
         Height          =   495
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Back"
         DisabledPicture =   "Form6.frx":1A9B5
         Height          =   495
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   10
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
         ItemData        =   "Form6.frx":35298
         Left            =   2400
         List            =   "Form6.frx":352A8
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   7080
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
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   6
         Top             =   4800
         WhatsThisHelpID =   1
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   8880
         MaxLength       =   20
         TabIndex        =   5
         Top             =   6240
         WhatsThisHelpID =   1
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   5280
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lecturer Name"
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
         Index           =   18
         Left            =   240
         TabIndex        =   19
         Top             =   4200
         Width           =   2055
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lecturer ID"
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
         Index           =   10
         Left            =   720
         TabIndex        =   18
         Top             =   4800
         Width           =   1635
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faculty"
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
         Left            =   720
         TabIndex        =   17
         Top             =   5400
         Width           =   1845
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   6000
         Width           =   2145
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Campus/Center"
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
         Left            =   6480
         TabIndex        =   15
         Top             =   4320
         Width           =   2325
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulding"
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
         Index           =   4
         Left            =   6840
         TabIndex        =   14
         Top             =   5040
         Width           =   2415
         WordWrap        =   -1  'True
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
         Index           =   5
         Left            =   6720
         TabIndex        =   13
         Top             =   5760
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rank"
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
         Left            =   6840
         TabIndex        =   12
         Top             =   6360
         Width           =   2895
         WordWrap        =   -1  'True
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
      ItemData        =   "Form6.frx":352E5
      Left            =   5160
      List            =   "Form6.frx":352E7
      TabIndex        =   1
      Top             =   4680
      Width           =   3255
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
         TabIndex        =   32
         Top             =   6960
         Width           =   1455
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
Attribute VB_Name = "manage_lec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataGrid1_Click()





End Sub

Private Sub Combo6_Change()
If Combo6.Text = "Professor" Or Combo6.Text = "Assistant Professor" Or Combo6.Text = "Senior Lecturer(HG)" Or Combo6.Text = "Senior Lecturer" Or Combo6.Text = "Lecturer" Or Combo6.Text = "Assistant Lecturer" Then


Text6.Text = ""
If Combo6.Text = "Professor" Then

Text6.Text = 1

ElseIf Combo6.Text = "Assistant Professor" Then

Text6.Text = 2

ElseIf Combo6.Text = "Senior Lecturer(HG)" Then

Text6.Text = 3

ElseIf Combo6.Text = "Senior Lecturer" Then

Text6.Text = 4

ElseIf Combo6.Text = "Lecturer" Then

Text6.Text = 5

ElseIf Combo6.Text = "Assistant Lecturer" Then

Text6.Text = 6

End If




Else
MsgBox "Cant edit", vbInformation
Combo6.SetFocus
End If
Text1.Text = Text6.Text + "." + Text8.Text
End Sub

Private Sub Combo6_Click()
Combo6_Change
End Sub

Private Sub Combo6_GotFocus()
Combo6_Change
End Sub

Private Sub Combo6_LostFocus()
Combo6_Change
End Sub

Private Sub Command2_Click()
If Text3.Text = "" Then
MsgBox "Enter the Lecture Name", vbInformation
Text3.SetFocus
Exit Sub
End If

If Text8.Text = "" Then
MsgBox "Enter the Lecture ID", vbInformation
Text8.SetFocus
Exit Sub
End If

If Combo1.Text = "" Then
MsgBox "Enter the Faculty", vbInformation
Combo1.SetFocus
Exit Sub
End If

If Combo3.Text = "" Then
MsgBox "Enter the Department", vbInformation
Combo3.SetFocus
Exit Sub
End If

If Combo4.Text = "" Then
MsgBox "Enter the Campus/Center", vbInformation
Combo4.SetFocus
Exit Sub
End If

If Combo5.Text = "" Then
MsgBox "Enter the Bulding", vbInformation
Combo5.SetFocus
Exit Sub
End If

If Combo6.Text = "" Then
MsgBox "Enter the Level", vbInformation
Combo6.SetFocus
Exit Sub
End If

If Text1.Text = "" Then
MsgBox "Enter the Rank ", vbInformation
Text1.SetFocus
Exit Sub
End If



sql = "select * from Lecturer where lec_id='" + Text8.Text + "'"
Set dataset = mddata(sql)

With dataset

.Update
!lec_name = Text3.Text
!lec_id = Text8.Text
!faculty = Combo1.Text
!department = Combo3.Text
!campus_center = Combo4.Text
!cat = Combo6.Text
!Level = Text4.Text
!rank = Text1.Text
!bulding = Combo5.Text
.Update

MsgBox "Successfully!", vbInformation

Unload Me
Me.Show
End With



End Sub

Private Sub Command3_Click()
If Not Text5.Text = "" Then
sql = "delete Lecturer where ID='" + Text5.Text + "'"
Set dataset = mddata(sql)
MsgBox "Done!!", vbInformation
Unload Me
Me.Show

Else
MsgBox "Select Tag!", vbInformation
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
lec_managment.Show
End Sub

Private Sub Form_Load()
sql = "select * from Lecturer"
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
End Sub

Private Sub MSHFlexGrid3_Click()
MSHFlexGrid3.Col = 1
Text5.Text = MSHFlexGrid3.Text





End Sub

Private Sub Text2_Change()

If Not Text2.Text = "" Then
sql = "select * from Lecturer where lec_id like '%" + Text2.Text + "%'"
Set dataset = mddata(sql)
With dataset

Set MSHFlexGrid3.DataSource = dataset


End With
End If
End Sub

Private Sub Text2_GotFocus()
Text2.Text = ""

Text2.FontSize = 12
Text2.FontName = "Arial"
End Sub

Private Sub Text5_Change()


sql = "select * from Lecturer where ID = '" + Text5.Text + "'"
Set dataset = mddata(sql)
With dataset

If dataset.RecordCount > 0 Then
Text3.Text = !lec_name
Text8.Text = !lec_id
Combo1.Text = !faculty
Combo3.Text = !department
Combo4.Text = !campus_center
Combo5.Text = !bulding
Combo6.Text = !cat
Text1.Text = !rank
Text4.Text = !rank

End If


End With


End Sub

Private Sub Text8_Change()
Text1.Text = Text6.Text + "." + Text8.Text
End Sub
