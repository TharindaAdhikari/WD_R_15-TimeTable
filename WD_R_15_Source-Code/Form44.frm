VERSION 5.00
Begin VB.Form add_wor_days_hou 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Working days Hours"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11280
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   14535
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add Working days Hours"
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
         Left            =   960
         TabIndex        =   23
         Top             =   240
         Width           =   12255
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
         Height          =   735
         Left            =   10440
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   15855
      Begin VB.CommandButton Command1 
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
         Left            =   7320
         TabIndex        =   31
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Working days "
         Height          =   3735
         Left            =   5520
         TabIndex        =   13
         Top             =   2040
         Width           =   3015
         Begin VB.CheckBox Check8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Check1"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   840
            Width           =   255
         End
         Begin VB.CheckBox Check7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Check1"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1320
            Width           =   255
         End
         Begin VB.CheckBox Check6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Check1"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1800
            Width           =   255
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Check1"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   2280
            Width           =   255
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Check1"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   2760
            Width           =   255
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Check1"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   3240
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Check1"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label152 
            BackStyle       =   0  'Transparent
            Caption         =   "Sunday"
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
            Index           =   9
            Left            =   720
            TabIndex        =   30
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label Label152 
            BackStyle       =   0  'Transparent
            Caption         =   "Saturday"
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
            Index           =   8
            Left            =   720
            TabIndex        =   29
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label152 
            BackStyle       =   0  'Transparent
            Caption         =   "Friday"
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
            Index           =   7
            Left            =   720
            TabIndex        =   28
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label152 
            BackStyle       =   0  'Transparent
            Caption         =   "Thursday"
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
            Index           =   6
            Left            =   720
            TabIndex        =   27
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label152 
            BackStyle       =   0  'Transparent
            Caption         =   "Wednesday"
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
            Index           =   5
            Left            =   720
            TabIndex        =   26
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label152 
            BackStyle       =   0  'Transparent
            Caption         =   "Tuesday"
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
            Left            =   720
            TabIndex        =   25
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label152 
            BackStyle       =   0  'Transparent
            Caption         =   "Monday"
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
            Left            =   720
            TabIndex        =   24
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Working Time Per Day"
         Height          =   3735
         Left            =   1680
         TabIndex        =   8
         Top             =   2040
         Width           =   3375
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
            Height          =   390
            Left            =   1680
            TabIndex        =   34
            Top             =   1920
            WhatsThisHelpID =   1
            Width           =   1335
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
            Height          =   390
            Left            =   1680
            TabIndex        =   33
            Top             =   960
            WhatsThisHelpID =   1
            Width           =   1335
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label152 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hours"
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
            Left            =   0
            TabIndex        =   9
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "Form44.frx":0000
         Height          =   495
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Back"
         DisabledPicture =   "Form44.frx":1A8E3
         Height          =   495
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
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
         ItemData        =   "Form44.frx":351C6
         Left            =   3600
         List            =   "Form44.frx":351DF
         TabIndex        =   0
         Top             =   1320
         Width           =   2535
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
         Left            =   5760
         TabIndex        =   5
         Top             =   6000
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
         Left            =   4200
         TabIndex        =   4
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No of working days per Week"
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
         TabIndex        =   11
         Top             =   1320
         Width           =   3165
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
         TabIndex        =   10
         Top             =   6480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   0
      TabIndex        =   1
      Top             =   1200
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
         TabIndex        =   32
         Top             =   5760
         Visible         =   0   'False
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
      Left            =   3720
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "add_wor_days_hou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Combo1.Text = "" Then

MsgBox "Enter the working days", vbInformation
Combo1.SetFocus
Exit Sub
End If

If Combo4.Text = "" Then

MsgBox "Enter the Hours", vbInformation
Combo4.SetFocus
Exit Sub
End If

If Combo3.Text = "" Then

MsgBox "Enter the Minutes", vbInformation
Combo3.SetFocus
Exit Sub
End If


sql = "select * from working_days_hours"
Set dataset = mddata(sql)



With dataset


.Update
!No_Of_Working_Days = Combo1.Text
!Hours = Combo4.Text
!Minutes = Combo3.Text
!monday = Check1.Value
!tuesday = Check8.Value
!wendesday = Check7.Value
!thursday = Check6.Value
!friday = Check5.Value
!saturday = Check4.Value
!sunday = Check3.Value
.Update

MsgBox "Successfully!", vbInformation
End With


End Sub

Private Sub Command2_Click()


If Combo1.Text = "" Then

MsgBox "Enter the working days", vbInformation
Combo1.SetFocus
Exit Sub
End If
'
If Text1.Text = "" Then
MsgBox "Enter the Hours", vbInformation
Text1.SetFocus
Exit Sub
End If

If Text2.Text = "" Then
MsgBox "Enter the Minutes", vbInformation
Text2.SetFocus
Exit Sub
End If


If Text1.Text > 24 Then
MsgBox "Enter the proper value", vbInformation
Text2.SetFocus
Exit Sub
End If



If Text2.Text > 60 Then
MsgBox "Enter the proper value", vbInformation
Text2.SetFocus
Exit Sub
End If



sql = "select * from working_days_hours"
Set dataset = mddata(sql)



With dataset


If dataset.RecordCount = 0 Then

.AddNew
!No_Of_Working_Days = Combo1.Text
!Hours = Text1.Text
!Minutes = Text2.Text
!monday = Check1.Value
!tuesday = Check8.Value
!wendesday = Check7.Value
!thursday = Check6.Value
!friday = Check5.Value
!saturday = Check4.Value
!sunday = Check3.Value


.Update


MsgBox "Successfully!", vbInformation

Else

MsgBox "Record allready Exceede"
End If
End With
End Sub

Private Sub Command3_Click()

sql = "delete from working_days_hours "
Set dataset = mddata(sql)

Unload Me
Me.Show

End Sub

Private Sub Command4_Click()
Unload Me
Home.Show
End Sub
