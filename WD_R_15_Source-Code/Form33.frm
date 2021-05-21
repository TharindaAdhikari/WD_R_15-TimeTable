VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form manage_session 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Session"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14805
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   14805
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "Form33.frx":0000
      Left            =   4920
      List            =   "Form33.frx":0002
      TabIndex        =   35
      Top             =   6000
      Width           =   3255
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
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
      Left            =   13920
      MaxLength       =   20
      TabIndex        =   30
      Top             =   7560
      WhatsThisHelpID =   1
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   0
      TabIndex        =   23
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
         TabIndex        =   24
         Top             =   6960
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
      ItemData        =   "Form33.frx":0004
      Left            =   4920
      List            =   "Form33.frx":0006
      TabIndex        =   22
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   15855
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   3960
         TabIndex        =   33
         Top             =   6360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   2520
         TabIndex        =   32
         Top             =   6360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   840
         TabIndex        =   31
         Top             =   6360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Add Session"
         DisabledPicture =   "Form33.frx":0008
         Height          =   495
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text3 
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
         TabIndex        =   27
         Top             =   5640
         WhatsThisHelpID =   1
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   5280
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         Left            =   9840
         MaxLength       =   20
         TabIndex        =   14
         Top             =   6360
         WhatsThisHelpID =   1
         Width           =   615
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         ItemData        =   "Form33.frx":1A8EB
         Left            =   2400
         List            =   "Form33.frx":1A8ED
         TabIndex        =   11
         Top             =   5400
         Width           =   3255
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Form33.frx":1A8EF
         Height          =   495
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   11895
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   2535
            Left            =   120
            TabIndex        =   9
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
         TabIndex        =   7
         Text            =   "Serch Here........"
         Top             =   840
         Width           =   2295
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
         ItemData        =   "Form33.frx":351D2
         Left            =   8880
         List            =   "Form33.frx":351D4
         TabIndex        =   6
         Top             =   4320
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
         ItemData        =   "Form33.frx":351D6
         Left            =   8880
         List            =   "Form33.frx":351D8
         TabIndex        =   5
         Top             =   5040
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Lecturer2"
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
         Left            =   240
         TabIndex        =   34
         Top             =   4920
         Width           =   2145
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes"
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
         Left            =   10320
         TabIndex        =   29
         Top             =   6360
         Width           =   1275
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hrs"
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
         Left            =   9240
         TabIndex        =   21
         Top             =   6360
         Width           =   795
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Of Student"
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
         Left            =   6705
         TabIndex        =   20
         Top             =   5760
         Width           =   2805
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Subject"
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
         Left            =   6825
         TabIndex        =   19
         Top             =   5040
         Width           =   2445
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Group"
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
         Left            =   6465
         TabIndex        =   18
         Top             =   4320
         Width           =   2355
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Tag"
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
         Left            =   720
         TabIndex        =   17
         Top             =   5520
         Width           =   1665
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Lecturer"
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
         Left            =   210
         TabIndex        =   16
         Top             =   4200
         Width           =   2115
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15855
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Session"
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
         TabIndex        =   1
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
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   5760
      Width           =   1455
   End
End
Attribute VB_Name = "manage_session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_GotFocus()
Combo1.Clear
sql = "select * from Tag"
Set Dataset2 = mddata(sql)

With Dataset2
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo1.AddItem Dataset2!Tag_Name
.MoveNext
Wend

End If
End With

End Sub


Private Sub Combo2_GotFocus()
Combo2.Clear
sql = "select * from Lecturer"
Set dataset = mddata(sql)

With dataset
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo2.AddItem dataset!lec_name
.MoveNext
Wend

End If
End With



End Sub

Private Sub Combo3_GotFocus()
Combo3.Clear
sql = "select * from Lecturer"
Set dataset = mddata(sql)

With dataset
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo3.AddItem dataset!lec_name
.MoveNext
Wend

End If
End With

End Sub


Private Sub Combo4_GotFocus()
Combo4.Clear
sql = "select * from Student_group"
Set dataset3 = mddata(sql)

With dataset3
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo4.AddItem dataset3!group_id
.MoveNext
Wend

End If
End With

End Sub


Private Sub Combo5_GotFocus()
Combo5.Clear
sql = "select * from Subject"
Set dataset4 = mddata(sql)

With dataset4
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo5.AddItem dataset4!Sub_name
.MoveNext
Wend

End If
End With
End Sub


Private Sub Command1_Click()
Unload Me
Add_session.Show
End Sub

Private Sub Command2_Click()

If Text5.Text = "" Then
MsgBox "Select the Session", vbInformation
Combo4.SetFocus
Exit Sub
End If

If val(Text1.Text) = 0 And val(Text7.Text) = 0 Then
MsgBox "error!", vbInformation
Text1.SetFocus
Exit Sub
End If

Text8.Text = val(Text1.Text)
Text9.Text = val(Text7.Text) / 30
Text10.Text = ""
Text10.Text = val(Text8.Text) + val(Text9.Text)

sql = "select * from session where ID='" + Text5.Text + "'"
Set dataset = mddata(sql)

With dataset

.Update
!lecturer = Combo2.Text

If Not Combo3.Text = "" Then
!lecturer2 = Combo3.Text
Else
!lecturer2 = ""
End If
!Tag = Combo1.Text
!Group = Combo5.Text
!subject = Combo4.Text
!nostudent = val(Text3.Text)
!duration_hours = val(Text1.Text)
!duration_minuts = val(Text7.Text)
If IsNull(!not_ava_time) = False Then
!val = val(Text10.Text) + 1
Else
!val = val(Text10.Text)
End If


.Update

MsgBox "Successfully!", vbInformation

Unload Me
Me.Show
End With

End Sub

Private Sub Command3_Click()
If Not Text5.Text = "" Then
sql = "delete session where ID='" + Text5.Text + "'"
Set dataset = mddata(sql)
MsgBox "Done!!", vbInformation
Unload Me
Me.Show

Else
MsgBox "Select session!", vbInformation
MSHFlexGrid3.SetFocus
Exit Sub

End If
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command4_Click()
Unload Me
Home.Show
End Sub

Private Sub Form_Load()

sql = "select * from session "
Set dataset = mddata(sql)
With dataset

Set MSHFlexGrid3.DataSource = dataset


End With

End Sub


Private Sub MSHFlexGrid3_Click()
MSHFlexGrid3.Col = 1
Text5.Text = MSHFlexGrid3.Text
End Sub

Private Sub Text1_LostFocus()


Text1.Text = Replace(Text1.Text, ".", "")
Text1.Text = Replace(Text1.Text, "'", "")

End Sub

Private Sub Text2_Change()

If Not Text2.Text = "" Then
sql = "select * from session where ID like '%" + Text2.Text + "%'"
Set dataset = mddata(sql)
With dataset

Set MSHFlexGrid3.DataSource = dataset


End With
End If

End Sub

Private Sub Text2_GotFocus()
Text2.Text = ""
End Sub

Private Sub Text5_Change()
sql = "select * from session where ID = '" + Text5.Text + "'"
Set dataset = mddata(sql)
With dataset

If dataset.RecordCount > 0 Then
Combo2.Text = !lecturer
If IsNull(!lecturer2) = False Then
Combo3.Text = !lecturer2
End If

Combo1.Text = !Tag
Combo5.Text = !Group
Combo4.Text = !subject
Text3.Text = !nostudent
Text1.Text = !duration_hours
Text7.Text = !duration_minuts


End If


End With
End Sub

Private Sub Text7_LostFocus()

If Not val(Text7.Text) = 30 And Not val(Text7.Text) = 0 Then
MsgBox "Error!", vbInformation
Text7.Text = 0
Text7.SetFocus
Exit Sub
End If




End Sub
