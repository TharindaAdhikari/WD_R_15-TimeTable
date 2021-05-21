VERSION 5.00
Begin VB.Form manage_session_room 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Session Room"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12300
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   12300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   2280
      TabIndex        =   4
      Top             =   1080
      Width           =   10695
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Back"
         DisabledPicture =   "manage_session_room.frx":0000
         Height          =   495
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   13
         Top             =   3000
         Width           =   3135
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
         ItemData        =   "manage_session_room.frx":1A8E3
         Left            =   2760
         List            =   "manage_session_room.frx":1A8E5
         TabIndex        =   9
         Top             =   2160
         WhatsThisHelpID =   2
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
         Left            =   3960
         TabIndex        =   8
         Top             =   3600
         WhatsThisHelpID =   6
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
         Left            =   2040
         TabIndex        =   7
         Top             =   3600
         WhatsThisHelpID =   5
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Home"
         DisabledPicture =   "manage_session_room.frx":1A8E7
         Height          =   495
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   1095
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
         ItemData        =   "manage_session_room.frx":351CA
         Left            =   2760
         List            =   "manage_session_room.frx":351CC
         TabIndex        =   5
         Top             =   1320
         WhatsThisHelpID =   2
         Width           =   2535
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select Room"
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
         TabIndex        =   12
         Top             =   2280
         Width           =   1455
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
         Left            =   8280
         TabIndex        =   11
         Top             =   6600
         Width           =   1455
      End
      Begin VB.Label Label152 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Session"
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
         Left            =   0
         TabIndex        =   10
         Top             =   1320
         Width           =   2985
         WordWrap        =   -1  'True
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
      TabIndex        =   3
      Top             =   840
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
         Left            =   360
         TabIndex        =   15
         Top             =   5520
         Width           =   1455
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H008B8B00&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Session Room"
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
         Left            =   1680
         TabIndex        =   2
         Top             =   120
         Width           =   11775
      End
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
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "manage_session_room"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo4_Change()
Text1.Text = Combo4.Text
End Sub

Private Sub Combo4_Click()
Text1.Text = Combo4.Text
End Sub

Private Sub Command2_Click()

If Combo4.Text = "" Then
MsgBox "Error!", vbInformation
Combo4.SetFocus
Exit Sub
End If

If Combo1.Text = "" Then
MsgBox "Error!", vbInformation
Combo1.SetFocus
Exit Sub
End If

sql = "select * from session where id='" + Combo4.Text + "'"
Set dataset3 = mddata(sql)

With dataset3
If .RecordCount > 0 Then
.Update
'!Session = Combo4.Text
!room = Combo1.Text
.Update
Else
MsgBox "Select Session!", vbInformation
Exit Sub
End If

End With

MsgBox "Done", vbInformation
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
Form4.Show
End Sub

Private Sub Form_Load()


sql = "select * from session order by ID"
Set dataset = mddata(sql)

With dataset
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo4.AddItem dataset!id
.MoveNext
Wend

End If
End With

sql = "select * from Loacation"
Set dataset = mddata(sql)

With dataset
If .RecordCount > 0 Then
.MoveFirst
While .EOF = False
Combo1.AddItem dataset!Room_Name
.MoveNext
Wend

End If
End With

End Sub

