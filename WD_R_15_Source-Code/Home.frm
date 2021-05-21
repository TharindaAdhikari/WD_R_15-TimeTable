VERSION 5.00
Begin VB.Form Home 
   BorderStyle     =   0  'None
   Caption         =   "Home"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13620
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   8415
   ScaleWidth      =   13620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      Begin VB.Label Label1 
         BackColor       =   &H008B8B00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   15
         Left            =   1080
         TabIndex        =   10
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   36
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   12600
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time Table Managment"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   39.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   1320
         TabIndex        =   1
         Top             =   0
         Width           =   11775
      End
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00404000&
      Caption         =   "Manage Session Room"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000B&
      Height          =   9735
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   16455
      Begin VB.CommandButton Command4 
         BackColor       =   &H00404000&
         Caption         =   "Session And Not Available Time"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00404000&
         Caption         =   "Manage Session"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00404000&
         Caption         =   "Genarate Time Table"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5880
         Width           =   2175
      End
      Begin VB.CommandButton Bill 
         BackColor       =   &H00404000&
         Caption         =   "Home"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H008B8B00&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   7455
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2535
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   960
         Top             =   7320
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Statistic"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   10560
         TabIndex        =   19
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Subjects"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   11160
         TabIndex        =   18
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Lecturer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   8520
         TabIndex        =   17
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   7800
         TabIndex        =   15
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Tags"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   6120
         TabIndex        =   14
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Add Working days Hours"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4800
         TabIndex        =   13
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Student Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3600
         TabIndex        =   12
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
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
         Left            =   12000
         TabIndex        =   11
         Top             =   6480
         Width           =   1455
      End
      Begin VB.Shape Shape7 
         Height          =   1575
         Left            =   11040
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Shape Shape6 
         Height          =   1455
         Left            =   10320
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Shape Shape5 
         Height          =   1455
         Left            =   7560
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Shape Shape4 
         Height          =   1455
         Left            =   4680
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Shape Shape3 
         Height          =   1575
         Left            =   8400
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Shape Shape2 
         Height          =   1575
         Left            =   6000
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         Height          =   1575
         Left            =   3480
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   2055
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tags"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command3_Click()

End Sub

Private Sub Command1_Click()
Timetable.Show
End Sub

Private Sub Command2_Click()
manage_session.Show
End Sub

Private Sub Command4_Click()
Session_And_Not_Time_Manage.Show
End Sub

Private Sub Command5_Click()
'Unload Me
Form4.Show
'manage_session_room.Show
End Sub

Private Sub Form_Activate()
' Timer1_Timer
End Sub

Private Sub Label10_Click()
Unload Me
Statictis.Show
End Sub

Private Sub Label3_Click()
Unload Me
Student_Group_Managment.Show
End Sub

Private Sub Label4_Click()
Unload Me
add_wor_days_hou.Show
End Sub

Private Sub Label5_Click()
Unload Me
Tag_Managment.Show
End Sub

Private Sub Label57_Click()


Select Case MsgBox("Are you sure you want to Exit ?", vbYesNo)
Case vbYes
Unload Me
End
Case vbNo

End Select

End Sub



Private Sub Label6_Click()
Unload Me
Location_Managment.Show
End Sub

Private Sub Label8_Click()
Unload Me
lec_managment.Show
End Sub

Private Sub Label9_Click()
Unload Me
Subjects_Managment.Show
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Format(Now, "h:m:ss AMPM")
End Sub
