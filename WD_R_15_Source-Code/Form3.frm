VERSION 5.00
Begin VB.Form add_lecturer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add_lecturer"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11235
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "Form3.frx":0000
      Left            =   4680
      List            =   "Form3.frx":0002
      TabIndex        =   18
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
      Left            =   -360
      TabIndex        =   12
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
         TabIndex        =   29
         Top             =   7200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   15855
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   7080
         TabIndex        =   28
         Top             =   4440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text3 
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
         Left            =   2640
         TabIndex        =   25
         Top             =   6360
         WhatsThisHelpID =   1
         Width           =   2535
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
         ItemData        =   "Form3.frx":0004
         Left            =   2640
         List            =   "Form3.frx":001A
         TabIndex        =   24
         Top             =   5640
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
         ItemData        =   "Form3.frx":0082
         Left            =   2640
         List            =   "Form3.frx":0084
         TabIndex        =   22
         Top             =   4920
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
         ItemData        =   "Form3.frx":0086
         Left            =   2640
         List            =   "Form3.frx":009C
         TabIndex        =   20
         Top             =   4200
         Width           =   3255
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
         TabIndex        =   15
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
         Left            =   5520
         TabIndex        =   9
         Top             =   7200
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
         Left            =   7200
         TabIndex        =   8
         Top             =   7200
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
         ItemData        =   "Form3.frx":00D2
         Left            =   2640
         List            =   "Form3.frx":00E2
         TabIndex        =   7
         Top             =   2760
         Width           =   3255
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Back"
         DisabledPicture =   "Form3.frx":011F
         Height          =   495
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Form3.frx":1AA02
         Height          =   495
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Genarate Rank"
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
         Left            =   3840
         TabIndex        =   4
         Top             =   7200
         Width           =   1335
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
         Left            =   420
         TabIndex        =   27
         Top             =   6360
         Width           =   2895
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
         Left            =   480
         TabIndex        =   23
         Top             =   5640
         Width           =   2775
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
         Left            =   480
         TabIndex        =   21
         Top             =   4920
         Width           =   2535
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
         Left            =   240
         TabIndex        =   19
         Top             =   4200
         Width           =   2325
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
         Left            =   480
         TabIndex        =   17
         Top             =   3480
         Width           =   2145
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
         Left            =   840
         TabIndex        =   16
         Top             =   2880
         Width           =   1845
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
         TabIndex        =   14
         Top             =   2160
         Width           =   1635
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
         TabIndex        =   11
         Top             =   6480
         Width           =   1455
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
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   2055
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14535
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add New Lecturer"
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
         Left            =   600
         TabIndex        =   2
         Top             =   240
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
      Left            =   3720
      TabIndex        =   13
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "add_lecturer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo5_LostFocus()
If Combo5.Text = "Professor" Or Combo5.Text = "Assistant Professor" Or Combo5.Text = "Senior Lecturer(HG)" Or Combo5.Text = "Senior Lecturer" Or Combo5.Text = "Lecturer" Or Combo5.Text = "Assistant Lecturer" Then


Text4.Text = ""
If Combo5.Text = "Professor" Then

Text4.Text = 1

ElseIf Combo5.Text = "Assistant Professor" Then

Text4.Text = 2

ElseIf Combo5.Text = "Senior Lecturer(HG)" Then

Text4.Text = 3

ElseIf Combo5.Text = "Senior Lecturer" Then

Text4.Text = 4

ElseIf Combo5.Text = "Lecturer" Then

Text4.Text = 5

ElseIf Combo5.Text = "Assistant Lecturer" Then

Text4.Text = 6

End If




Else
MsgBox "Cant edit", vbInformation
Combo5.SetFocus
End If
End Sub

Private Sub Command1_Click()

If Text4.Text = "" Then
MsgBox "select level", vbInformation
Combo5.SetFocus
Exit Sub
End If

If Text1.Text = "" Then
MsgBox "select Lecture ID", vbInformation
Text1.SetFocus
Exit Sub
End If

Text3.Text = Text4.Text + "." + Text1.Text
End Sub

Private Sub Command2_Click()

If Text8.Text = "" Then
MsgBox "Enter the Lecture Name", vbInformation
Text8.SetFocus
Exit Sub
End If

If Text1.Text = "" Then
MsgBox "Enter the Lecture ID", vbInformation
Text1.SetFocus
Exit Sub
End If

If Combo1.Text = "" Then
MsgBox "Enter the Faculty", vbInformation
Combo1.SetFocus
Exit Sub
End If

If Combo2.Text = "" Then
MsgBox "Enter the Department", vbInformation
Combo2.SetFocus
Exit Sub
End If

If Combo3.Text = "" Then
MsgBox "Enter the Campus/Center", vbInformation
Combo3.SetFocus
Exit Sub
End If

If Combo4.Text = "" Then
MsgBox "Enter the Bulding", vbInformation
Combo4.SetFocus
Exit Sub
End If

If Combo5.Text = "" Then
MsgBox "Enter the Level", vbInformation
Combo5.SetFocus
Exit Sub
End If

If Text3.Text = "" Then
MsgBox "Enter the Rank or Genarate Rank", vbInformation
Text3.SetFocus
Exit Sub
End If


sql = "select * from Lecturer where lec_id='" + Text1.Text + "'"
Set dataset1 = mddata(sql)

With dataset1

If dataset1.RecordCount = 0 Then
'If dataset1.recordecount = 0 Then



sql = "select * from Lecturer"
Set dataset = mddata(sql)

With dataset

.AddNew
!lec_name = Text8.Text
!lec_id = Text1.Text
!faculty = Combo1.Text
!department = Combo2.Text
!campus_center = Combo3.Text
!cat = Combo5.Text
!Level = Text4.Text
!rank = Text3.Text
!bulding = Combo4.Text
.Update
MsgBox "Successfully!", vbInformation

Unload Me
Me.Show
End With


Else
MsgBox "Lecturer id allready exceed!", vbInformation
Text1.SetFocus
Exit Sub

End If

End With
End Sub

Private Sub Command3_Click()
Unload Me
Me.Show
End Sub

Private Sub Command4_Click()
Unload Me
Home.Show
End Sub

Private Sub Label57_Click()
Unload Me
Home.Show
End Sub

Private Sub Command5_Click()
Unload Me
lec_managment.Show
End Sub


